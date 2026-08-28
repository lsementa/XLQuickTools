using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    public partial class TextConvertForm : Form
    {
        private readonly Excel.Application _excelApp;

        private const string LocaleUs = "United States";
        private const string LocaleGlobal = "Global";

        // Category names
        private const string CatDate = "Date";
        private const string CatPhone = "Phone Number";
        private const string CatZip = "Zip Code";
        private const string CatSsn = "Social Security Number";
        private const string CatEin = "EIN / Tax ID";
        private const string CatEmail = "Email";
        private const string CatTime = "Time / Duration";
        private const string CatRegion = "State / Country";

        // Date formats that .NET cannot express, handled by FormatSpecialDate
        private const string FmtQuarterFirst = "Q# yyyy";
        private const string FmtQuarterLast = "yyyy-Q#";
        private const string FmtIsoWeek = "yyyy-Www";
        private const string FmtOrdinalDay = "yyyy-DDD";

        private static readonly string[] SpecialDateFormats =
        {
            FmtQuarterFirst,
            FmtQuarterLast,
            FmtIsoWeek,
            FmtOrdinalDay
        };

        // MDY (United States)
        private static readonly object[] UsOutputFormats =
        {
            "yyyy-MM-dd",
            "M/d/yyyy",
            "MM/dd/yyyy",
            "MM-dd-yyyy",
            "MM.dd.yyyy",
            "M/d/yy",
            "MM/dd/yy",
            "MMddyyyy",
            "yyyyMMdd",
            "MMM-d-yyyy",
            "d-MMM-yyyy",
            "dd-MMM-yy",
            "MMM d, yyyy",
            "MMM dd, yyyy",
            "MMMM d, yyyy",
            "MMMM dd, yyyy",
            "dddd, MMMM d, yyyy",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "yyyy MMM dd",
            "MMM yyyy",
            "yyyy-MM",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss",
            "MM/dd/yyyy h:mm tt",
            FmtQuarterFirst,
            FmtQuarterLast,
            FmtIsoWeek,
            FmtOrdinalDay
        };

        // DMY (Global)
        private static readonly object[] GlobalOutputFormats =
        {
            "yyyy-MM-dd",
            "dd/MM/yyyy",
            "d/M/yyyy",
            "dd-MM-yyyy",
            "dd.MM.yyyy",
            "d.M.yyyy",
            "d/M/yy",
            "dd/MM/yy",
            "ddMMyyyy",
            "yyyyMMdd",
            "d-MMM-yyyy",
            "dd-MMM-yy",
            "d MMM yyyy",
            "dd MMM yyyy",
            "d MMMM yyyy",
            "dd MMMM yyyy",
            "dddd, d MMMM yyyy",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "yyyy MMM dd",
            "MMM yyyy",
            "yyyy-MM",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss",
            "dd/MM/yyyy HH:mm",
            FmtQuarterFirst,
            FmtQuarterLast,
            FmtIsoWeek,
            FmtOrdinalDay
        };

        // Input formats used for parsing. FormatDateToString and ConvertToExcelSerial
        private static readonly string[] UsInputFormats =
        {
            // ISO / year-first (locale neutral)
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "yyyy MMM dd",
            "yyyy-MM-dd HH:mm",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-ddTHH:mm",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-ddTHH:mm:ssZ",
            "yyyy-MM-ddTHH:mm:ss.fff",
            "yyyy-MM-ddTHH:mm:ss.fffZ",
            "yyyy-MM-ddTHH:mm:ss.ffffffZ",

            // Month-name forms (locale neutral)
            "MMM-d-yyyy",
            "MMM-dd-yyyy",
            "d-MMM-yyyy",
            "dd-MMM-yyyy",
            "d-MMM-yy",
            "dd-MMM-yy",
            "MMM d, yyyy",
            "MMM dd, yyyy",
            "MMMM d, yyyy",
            "MMMM dd, yyyy",
            "dddd, MMMM d, yyyy",
            "MMM d, yyyy h:mm tt",
            "MMM dd, yyyy HH:mm",
            "MMMM d, yyyy h:mm tt",
            "MMMM dd, yyyy h:mm",

            // Numeric MDY
            "M/d/yyyy",
            "MM/dd/yyyy",
            "M-d-yyyy",
            "MM-dd-yyyy",
            "M.d.yyyy",
            "MM.dd.yyyy",
            "M/d/yyyy H:mm",
            "MM/dd/yyyy H:mm",
            "M/d/yyyy h:mm tt",
            "MM/dd/yyyy h:mm tt",
            "M/d/yyyy HH:mm:ss",
            "MM/dd/yyyy HH:mm:ss",

            // All-digit: locale form first, then ISO basic
            "MMddyyyy",
            "yyyyMMdd",

            // Month/year keys
            "MMM yyyy",
            "MMMM yyyy",
            "yyyy-MM"
        };

        private static readonly string[] GlobalInputFormats =
        {
            // ISO / year-first (locale neutral)
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "yyyy MMM dd",
            "yyyy-MM-dd HH:mm",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-ddTHH:mm",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-ddTHH:mm:ssZ",
            "yyyy-MM-ddTHH:mm:ss.fff",
            "yyyy-MM-ddTHH:mm:ss.fffZ",
            "yyyy-MM-ddTHH:mm:ss.ffffffZ",

            // Month-name forms (locale neutral)
            "d-MMM-yyyy",
            "dd-MMM-yyyy",
            "d-MMM-yy",
            "dd-MMM-yy",
            "MMM-d-yyyy",
            "MMM-dd-yyyy",
            "d MMM yyyy",
            "dd MMM yyyy",
            "d MMMM yyyy",
            "dd MMMM yyyy",
            "dddd, d MMMM yyyy",
            "d MMM yyyy HH:mm",
            "dd MMM yyyy HH:mm",
            "d MMMM yyyy h:mm tt",
            "dd MMMM yyyy h:mm",

            // Numeric DMY
            "d/M/yyyy",
            "dd/MM/yyyy",
            "d-M-yyyy",
            "dd-MM-yyyy",
            "d.M.yyyy",
            "dd.MM.yyyy",
            "d/M/yyyy H:mm",
            "dd/MM/yyyy H:mm",
            "d/M/yyyy h:mm tt",
            "dd/MM/yyyy h:mm tt",
            "d/M/yyyy HH:mm:ss",
            "dd/MM/yyyy HH:mm:ss",

            // All-digit: locale form first, then ISO basic
            "ddMMyyyy",
            "yyyyMMdd",

            // Month/year keys
            "MMM yyyy",
            "MMMM yyyy",
            "yyyy-MM"
        };

        // "CODE|Name" pairs. States, DC and the US territories
        private static readonly string[] StateTable =
        {
            "AL|Alabama", "AK|Alaska", "AZ|Arizona", "AR|Arkansas", "CA|California",
            "CO|Colorado", "CT|Connecticut", "DE|Delaware", "FL|Florida", "GA|Georgia",
            "HI|Hawaii", "ID|Idaho", "IL|Illinois", "IN|Indiana", "IA|Iowa",
            "KS|Kansas", "KY|Kentucky", "LA|Louisiana", "ME|Maine", "MD|Maryland",
            "MA|Massachusetts", "MI|Michigan", "MN|Minnesota", "MS|Mississippi", "MO|Missouri",
            "MT|Montana", "NE|Nebraska", "NV|Nevada", "NH|New Hampshire", "NJ|New Jersey",
            "NM|New Mexico", "NY|New York", "NC|North Carolina", "ND|North Dakota", "OH|Ohio",
            "OK|Oklahoma", "OR|Oregon", "PA|Pennsylvania", "RI|Rhode Island", "SC|South Carolina",
            "SD|South Dakota", "TN|Tennessee", "TX|Texas", "UT|Utah", "VT|Vermont",
            "VA|Virginia", "WA|Washington", "WV|West Virginia", "WI|Wisconsin", "WY|Wyoming",
            "DC|District of Columbia", "PR|Puerto Rico", "VI|Virgin Islands", "GU|Guam",
            "AS|American Samoa", "MP|Northern Mariana Islands"
        };

        // Names RegionInfo either spells differently or does not carry at all
        private static readonly string[] CountryAliasTable =
        {
            "US|USA", "US|U.S.", "US|U.S.A.", "US|United States of America",
            "GB|UK", "GB|U.K.", "GB|Great Britain", "GB|England", "GB|Scotland",
            "GB|Wales", "GB|Northern Ireland",
            "KR|South Korea", "KR|Republic of Korea", "KP|North Korea",
            "RU|Russia", "RU|Russian Federation",
            "VN|Vietnam", "VN|Viet Nam",
            "IR|Iran", "SY|Syria", "TZ|Tanzania", "BO|Bolivia", "VE|Venezuela",
            "CZ|Czechia", "CZ|Czech Republic",
            "CI|Ivory Coast", "CI|Cote d'Ivoire",
            "CV|Cape Verde", "CV|Cabo Verde",
            "MM|Myanmar", "MM|Burma", "LA|Laos", "MD|Moldova",
            "MK|Macedonia", "MK|North Macedonia",
            "NL|Holland", "NL|The Netherlands",
            "AE|UAE", "AE|United Arab Emirates",
            "TW|Taiwan", "HK|Hong Kong", "MO|Macau",
            "PS|Palestine", "VA|Vatican City", "VA|Holy See",
            "TR|Turkey", "TR|Turkiye", "SZ|Eswatini", "SZ|Swaziland",
            "CD|Democratic Republic of the Congo", "CD|DRC", "CG|Republic of the Congo",
            "AQ|Antarctica", "CU|Cuba", "SD|Sudan", "SS|South Sudan"
        };

        private static Dictionary<string, string> _stateNameToCode;
        private static Dictionary<string, string> _stateCodeToName;
        private static Dictionary<string, string> _countryNameToCode;
        private static Dictionary<string, string> _countryCodeToName;

        public TextConvertForm(Excel.Application excelApp)
        {
            InitializeComponent();
            _excelApp = excelApp;
        }

        // On Load
        private void TextConvertForm_Load(object sender, EventArgs e)
        {
            // Populate the conversion type
            this.CbConvertType.Items.AddRange(new object[]
            {
                "Text",
                "Excel Format"
            });

            // Set the default conversion type
            this.CbConvertType.SelectedItem = "Text";

            // Populate the current locale
            this.CbCurrentLocale.Items.AddRange(new object[]
            {
                LocaleUs,
                LocaleGlobal
            });

            // Set the default locale
            this.CbCurrentLocale.SelectedItem = LocaleUs;

            // Populate the convert locale
            this.CbConvertLocale.Items.AddRange(new object[]
            {
                LocaleUs,
                LocaleGlobal
            });

            // Set the default locale
            this.CbConvertLocale.SelectedItem = LocaleUs;

            // Populate the category list with options
            this.CbCategory.Items.AddRange(new object[]
            {
                CatDate,
                CatPhone,
                CatZip,
                CatSsn,
                CatEin,
                CatEmail,
                CatTime,
                CatRegion
            });

            // Set the default category
            this.CbCategory.SelectedItem = CatDate;

            // Subscribe to the SelectedIndexChanged event
            this.CbCategory.SelectedIndexChanged += CbCategory_SelectedIndexChanged;

            // Populate the formats and the convert type for the default category
            UpdateFieldOptions();
        }

        // Category changed
        private void CbCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFieldOptions();
        }

        // Locale boxes only apply to dates, every other category turns them off
        private void SetLocaleEnabled(bool enabled)
        {
            if (!enabled)
            {
                this.CbCurrentLocale.Text = "";
                this.CbConvertLocale.Text = "";
            }

            this.CbCurrentLocale.Enabled = enabled;
            this.CbConvertLocale.Enabled = enabled;
            this.LblCurrentLocale.Enabled = enabled;
            this.LblFormatLocale.Enabled = enabled;
        }

        // Update all form options
        private void UpdateFieldOptions()
        {
            // Clear previous format options
            CbFormat.Items.Clear();
            TbExample.Text = string.Empty;

            // Get the selected locale and category
            string convertLocale = CbConvertLocale.SelectedItem?.ToString() ?? string.Empty;
            string currentLocale = CbCurrentLocale.SelectedItem?.ToString() ?? string.Empty;
            string category = CbCategory.SelectedItem?.ToString() ?? string.Empty;

            // Ensure category is not null or empty before proceeding
            if (string.IsNullOrEmpty(category)) return;

            // Populate CbFormat based on the selected item in CbCategory
            switch (category)
            {
                case CatDate:
                    if (convertLocale.Equals(LocaleGlobal))
                    {
                        // Populate DMY date formats
                        this.CbFormat.Items.AddRange(GlobalOutputFormats);
                    }
                    else
                    {
                        // Populate MDY date formats
                        this.CbFormat.Items.AddRange(UsOutputFormats);
                    }

                    if (currentLocale.Equals(LocaleGlobal))
                    {
                        // DMY Example
                        this.TbExample.Text = DateTime.Now.ToString("d/M/yyyy", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        this.TbExample.Text = DateTime.Now.ToString("M/d/yyyy", CultureInfo.InvariantCulture);
                    }

                    SetLocaleEnabled(true);

                    // If null from category change then set it
                    if (string.IsNullOrEmpty(this.CbCurrentLocale.Text))
                    {
                        this.CbCurrentLocale.SelectedItem = LocaleUs;
                    }
                    if (string.IsNullOrEmpty(this.CbConvertLocale.Text))
                    {
                        this.CbConvertLocale.SelectedItem = LocaleUs;
                    }

                    break;

                case CatPhone:
                    this.CbFormat.Items.AddRange(new object[]
                    {
                        "(###) ###-####",
                        "###/###-####",
                        "###-###-####",
                        "###.###.####",
                        "##########",
                        "+1 (###) ###-####",
                        "+1##########"
                    });
                    this.TbExample.Text = "6175550100";
                    this.CbConvertType.SelectedItem = "Text";
                    SetLocaleEnabled(false);
                    break;

                case CatZip:
                    this.CbFormat.Items.AddRange(new object[]
                    {
                        "#####",
                        "#####-####",
                        "#########",
                        "A1A 1A1",
                        "AA9A 9AA"
                    });
                    this.TbExample.Text = "02215-9997";
                    this.CbConvertType.SelectedItem = "Text";
                    SetLocaleEnabled(false);
                    break;

                case CatSsn:
                    this.CbFormat.Items.AddRange(new object[]
                    {
                        "###-##-####",
                        "#########"
                    });
                    this.TbExample.Text = "123456789";
                    this.CbConvertType.SelectedItem = "Text";
                    SetLocaleEnabled(false);
                    break;

                case CatEin:
                    this.CbFormat.Items.AddRange(new object[]
                    {
                        "##-#######",
                        "#########"
                    });
                    this.TbExample.Text = "123456789";
                    this.CbConvertType.SelectedItem = "Text";
                    SetLocaleEnabled(false);
                    break;

                case CatEmail:
                    this.CbFormat.Items.AddRange(new object[]
                    {
                        "Clean (extract + lowercase)",
                        "Extract address",
                        "Lowercase",
                        "User name only",
                        "Domain only"
                    });
                    this.TbExample.Text = "Test, John <jtest@bu.edu>";
                    this.CbConvertType.SelectedItem = "Text";
                    SetLocaleEnabled(false);
                    break;

                case CatTime:
                    this.CbFormat.Items.AddRange(new object[]
                    {
                        "HH:MM",
                        "HH:MM:SS",
                        "H:MM",
                        "Decimal hours",
                        "Total minutes"
                    });
                    this.TbExample.Text = "7:30";
                    this.CbConvertType.SelectedItem = "Text";
                    SetLocaleEnabled(false);
                    break;

                case CatRegion:
                    this.CbFormat.Items.AddRange(new object[]
                    {
                        "State name to code",
                        "State code to name",
                        "Country name to code",
                        "Country code to name"
                    });
                    this.TbExample.Text = "Massachusetts";
                    this.CbConvertType.SelectedItem = "Text";
                    SetLocaleEnabled(false);
                    break;

            }

            // Set a default format
            if (CbFormat.Items.Count > 0)
            {
                CbFormat.SelectedIndex = 0;
            }

            UpdateConvertTypeOptions();
        }

        // OK button
        private void TextConvertForm_Ok_Click(object sender, EventArgs e)
        {
            // Form values - bail out rather than throwing if the user cleared a box
            string category = CbCategory.SelectedItem?.ToString();
            string format = CbFormat.SelectedItem?.ToString();
            string convertType = CbConvertType.SelectedItem?.ToString() ?? "Text";
            string currentLocale = CbCurrentLocale.SelectedItem?.ToString() ?? LocaleUs;
            string convertLocale = CbConvertLocale.SelectedItem?.ToString() ?? LocaleUs;

            if (string.IsNullOrEmpty(category) || string.IsNullOrEmpty(format))
            {
                MessageBox.Show("Select a category and a format before converting.",
                    "XLQuickTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Never let an invalid pairing reach Excel's NumberFormat
            if (convertType == "Excel Format" && !SupportsExcelFormat(category, format))
            {
                convertType = "Text";
            }

            Excel.Range selectedRange = _excelApp.Selection as Excel.Range;
            if (selectedRange == null) return;

            // Create an instance of QTClipboard
            QTClipboard clipboard = QTClipboard.Instance;
            // Copy and store values
            clipboard.CopyAndStoreFormat(selectedRange);
            // Enable the Undo button
            Globals.Ribbons.Ribbon1.BtnUndo.Enabled = true;

            // Run
            ConvertValues(_excelApp, category, format, convertType, convertLocale, currentLocale);

            this.Close();
        }

        // Cancel button
        private void TextConvertForm_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Method to extract digits only
        private string DigitsOnly(string s)
        {
            string digitsOnly = new string(s.Where(char.IsDigit).ToArray());
            return digitsOnly;
        }

        // Letters and digits only, used by the non-US postal codes
        private string AlphaNumOnly(string s)
        {
            return new string(s.Where(char.IsLetterOrDigit).ToArray());
        }

        // Method to format a phone number based on the specified format
        private string FormatPhone(string phoneNumber, string format)
        {
            if (string.IsNullOrWhiteSpace(phoneNumber))
            {
                return string.Empty;
            }

            string text = phoneNumber.Trim();

            // Pull the extension off the end first, otherwise its digits get
            // folded into the number and the length test fails
            string extension = string.Empty;
            Match extMatch = Regex.Match(text,
                @"\b(?:x|ext|extn|extension)\.?\s*[:#]?\s*(\d{1,6})\s*$",
                RegexOptions.IgnoreCase);

            if (extMatch.Success)
            {
                extension = extMatch.Groups[1].Value;
                text = text.Substring(0, extMatch.Index);
            }

            string digitsOnly = DigitsOnly(text);
            bool wantsCountryCode = format.StartsWith("+1");

            // Normalize the leading country code to whatever the format wants
            if (digitsOnly.Length == 11 && digitsOnly.StartsWith("1") && !wantsCountryCode)
            {
                digitsOnly = digitsOnly.Substring(1);
            }
            else if (digitsOnly.Length == 10 && wantsCountryCode)
            {
                digitsOnly = "1" + digitsOnly;
            }

            string formatted = null;

            // 7-digit local numbers get a sensible form no matter which
            // 10-digit layout is selected
            if (digitsOnly.Length == 7)
            {
                formatted = $"{digitsOnly.Substring(0, 3)}-{digitsOnly.Substring(3, 4)}";
            }
            else if (digitsOnly.Length == 10 && !wantsCountryCode)
            {
                string a = digitsOnly.Substring(0, 3);
                string b = digitsOnly.Substring(3, 3);
                string c = digitsOnly.Substring(6, 4);

                switch (format)
                {
                    case "(###) ###-####":
                        formatted = $"({a}) {b}-{c}";
                        break;

                    case "###/###-####":
                        formatted = $"{a}/{b}-{c}";
                        break;

                    case "###-###-####":
                        formatted = $"{a}-{b}-{c}";
                        break;

                    case "###.###.####":
                        formatted = $"{a}.{b}.{c}";
                        break;

                    case "##########":
                        formatted = digitsOnly;
                        break;
                }
            }
            else if (digitsOnly.Length == 11 && digitsOnly.StartsWith("1") && wantsCountryCode)
            {
                string a = digitsOnly.Substring(1, 3);
                string b = digitsOnly.Substring(4, 3);
                string c = digitsOnly.Substring(7, 4);

                switch (format)
                {
                    case "+1 (###) ###-####":
                        formatted = $"+1 ({a}) {b}-{c}";
                        break;

                    case "+1##########":
                        formatted = "+" + digitsOnly;
                        break;
                }
            }

            // If the format doesn't match, return the original phone number
            if (formatted == null) return phoneNumber;

            return string.IsNullOrEmpty(extension) ? formatted : $"{formatted} x{extension}";
        }

        // Method to format ZIP code
        private string FormatZip(string zipCode, string format)
        {
            if (string.IsNullOrEmpty(zipCode)) return string.Empty;

            // Canadian and UK codes carry letters, so they never go through DigitsOnly
            if (format == "A1A 1A1")
            {
                string canada = AlphaNumOnly(zipCode).ToUpperInvariant();
                return canada.Length == 6
                    ? $"{canada.Substring(0, 3)} {canada.Substring(3, 3)}"
                    : zipCode;
            }

            if (format == "AA9A 9AA")
            {
                // UK outward code is 2-4 chars, the inward code is always 3
                string uk = AlphaNumOnly(zipCode).ToUpperInvariant();
                return uk.Length >= 5 && uk.Length <= 7
                    ? $"{uk.Substring(0, uk.Length - 3)} {uk.Substring(uk.Length - 3)}"
                    : zipCode;
            }

            // Extract digits only
            string digitsOnly = DigitsOnly(zipCode);

            // Restore the leading zero Excel drops on New England codes
            if (digitsOnly.Length == 4) digitsOnly = "0" + digitsOnly;
            if (digitsOnly.Length == 8) digitsOnly = "0" + digitsOnly;

            switch (format)
            {
                case "#####":
                    // Truncates ZIP+4 down to the base code
                    if (digitsOnly.Length == 5) return digitsOnly;
                    if (digitsOnly.Length == 9) return digitsOnly.Substring(0, 5);
                    break;

                case "#####-####":
                    if (digitsOnly.Length == 9)
                    {
                        return $"{digitsOnly.Substring(0, 5)}-{digitsOnly.Substring(5, 4)}";
                    }
                    if (digitsOnly.Length == 5) return digitsOnly;
                    break;

                case "#########":
                    if (digitsOnly.Length == 9 || digitsOnly.Length == 5) return digitsOnly;
                    break;
            }

            return zipCode; // Return original if no valid format
        }

        // Method to format a single SSN string
        private string FormatSSN(string ssn, string format)
        {
            if (string.IsNullOrEmpty(ssn)) return string.Empty;

            // Extract digits only
            string digitsOnly = DigitsOnly(ssn);

            // Restore the leading zero Excel drops
            if (digitsOnly.Length == 8) digitsOnly = "0" + digitsOnly;
            if (digitsOnly.Length != 9) return ssn;

            if (format == "#########") return digitsOnly;

            return $"{digitsOnly.Substring(0, 3)}-{digitsOnly.Substring(3, 2)}-{digitsOnly.Substring(5, 4)}";
        }

        // EIN and other 9-digit tax IDs, split 2-7 rather than 3-2-4
        private string FormatEIN(string ein, string format)
        {
            if (string.IsNullOrEmpty(ein)) return string.Empty;

            string digitsOnly = DigitsOnly(ein);

            if (digitsOnly.Length == 8) digitsOnly = "0" + digitsOnly;
            if (digitsOnly.Length != 9) return ein;

            if (format == "#########") return digitsOnly;

            return $"{digitsOnly.Substring(0, 2)}-{digitsOnly.Substring(2, 7)}";
        }

        // Pull a bare address out of "Test, John <jtest@bu.edu>" or "mailto:jtest@bu.edu"
        private string ExtractEmail(string value)
        {
            string text = value.Trim();

            Match angle = Regex.Match(text, @"<\s*([^<>\s]+@[^<>\s]+)\s*>");
            if (angle.Success) text = angle.Groups[1].Value;

            text = Regex.Replace(text, @"^\s*mailto:", string.Empty, RegexOptions.IgnoreCase);

            Match bare = Regex.Match(text, @"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}");
            return bare.Success ? bare.Value : text.Trim();
        }

        private string FormatEmail(string value, string format)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;

            switch (format)
            {
                case "Lowercase":
                    return value.Trim().ToLowerInvariant();

                case "Extract address":
                    return ExtractEmail(value);

                case "Clean (extract + lowercase)":
                    return ExtractEmail(value).ToLowerInvariant();

                case "User name only":
                    {
                        string address = ExtractEmail(value);
                        int at = address.IndexOf('@');
                        return at > 0 ? address.Substring(0, at).ToLowerInvariant() : value;
                    }

                case "Domain only":
                    {
                        string address = ExtractEmail(value);
                        int at = address.IndexOf('@');
                        return at >= 0 && at < address.Length - 1
                            ? address.Substring(at + 1).ToLowerInvariant()
                            : value;
                    }
            }

            return value;
        }

        // Parse a duration into hours. Handles clock text, "1h 30m", decimal
        // hours, and the fraction-of-a-day doubles Excel hands back
        private bool TryParseDuration(object value, out double totalHours)
        {
            totalHours = 0;
            if (value == null) return false;

            if (value is DateTime dt)
            {
                totalHours = dt.TimeOfDay.TotalHours;
                return true;
            }

            if (value is double d)
            {
                if (d >= 1000)
                {
                    // Full date serial, keep only the time portion
                    totalHours = (d - Math.Floor(d)) * 24.0;
                }
                else if (d < 1 && d > -1)
                {
                    // Excel stores a bare time as a fraction of a day
                    totalHours = d * 24.0;
                }
                else
                {
                    // A plain number typed by hand reads as decimal hours
                    totalHours = d;
                }
                return true;
            }

            string text = value.ToString().Trim();
            if (text.Length == 0) return false;

            bool negative = text.StartsWith("-");
            if (negative) text = text.Substring(1).Trim();

            if (text.Contains(":"))
            {
                string[] parts = text.Split(':');
                double hours = 0, minutes = 0, seconds = 0;

                if (!double.TryParse(parts[0], NumberStyles.Any, CultureInfo.InvariantCulture, out hours))
                    return false;
                if (parts.Length > 1 && !double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out minutes))
                    return false;
                if (parts.Length > 2 && !double.TryParse(parts[2], NumberStyles.Any, CultureInfo.InvariantCulture, out seconds))
                    return false;

                totalHours = hours + (minutes / 60.0) + (seconds / 3600.0);
                if (negative) totalHours = -totalHours;
                return true;
            }

            // "1h 30m" / "90m" / "2h"
            Match hm = Regex.Match(text, @"^(?:(\d+(?:\.\d+)?)\s*h)?\s*(?:(\d+(?:\.\d+)?)\s*m)?$",
                RegexOptions.IgnoreCase);
            if (hm.Success && (hm.Groups[1].Success || hm.Groups[2].Success))
            {
                double hours = hm.Groups[1].Success
                    ? double.Parse(hm.Groups[1].Value, CultureInfo.InvariantCulture) : 0;
                double minutes = hm.Groups[2].Success
                    ? double.Parse(hm.Groups[2].Value, CultureInfo.InvariantCulture) : 0;

                totalHours = hours + (minutes / 60.0);
                if (negative) totalHours = -totalHours;
                return true;
            }

            double plain;
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out plain))
            {
                totalHours = negative ? -plain : plain;
                return true;
            }

            return false;
        }

        private string FormatTime(object value, string format)
        {
            double totalHours;
            if (!TryParseDuration(value, out totalHours)) return value?.ToString() ?? string.Empty;

            bool negative = totalHours < 0;
            double abs = Math.Abs(totalHours);
            string sign = negative ? "-" : string.Empty;

            switch (format)
            {
                case "Decimal hours":
                    return (Math.Round(totalHours, 2)).ToString("0.##", CultureInfo.InvariantCulture);

                case "Total minutes":
                    return Math.Round(totalHours * 60.0).ToString("0", CultureInfo.InvariantCulture);
            }

            // Round to the second first so 1.99999 hours does not print as 1:59:60
            long totalSeconds = (long)Math.Round(abs * 3600.0);
            long hours = totalSeconds / 3600;
            long minutes = (totalSeconds % 3600) / 60;
            long seconds = totalSeconds % 60;

            switch (format)
            {
                case "HH:MM":
                    return $"{sign}{hours:00}:{minutes:00}";

                case "HH:MM:SS":
                    return $"{sign}{hours:00}:{minutes:00}:{seconds:00}";

                case "H:MM":
                    return $"{sign}{hours}:{minutes:00}";
            }

            return value.ToString();
        }

        private static void EnsureRegionMaps()
        {
            if (_stateNameToCode != null) return;

            var stateNameToCode = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var stateCodeToName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (string entry in StateTable)
            {
                string[] parts = entry.Split('|');
                stateCodeToName[parts[0]] = parts[1];
                stateNameToCode[parts[1]] = parts[0];
            }

            var countryNameToCode = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var countryCodeToName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Built from the installed cultures rather than a hard-coded ISO table
            foreach (CultureInfo culture in CultureInfo.GetCultures(CultureTypes.SpecificCultures))
            {
                try
                {
                    var region = new RegionInfo(culture.Name);
                    if (region.TwoLetterISORegionName.Length != 2) continue;

                    if (!countryCodeToName.ContainsKey(region.TwoLetterISORegionName))
                        countryCodeToName[region.TwoLetterISORegionName] = region.EnglishName;

                    if (!countryNameToCode.ContainsKey(region.EnglishName))
                        countryNameToCode[region.EnglishName] = region.TwoLetterISORegionName;
                }
                catch
                {
                    // Some cultures have no region, skip them
                }
            }

            // Aliases and the handful of codes no installed culture covers
            foreach (string entry in CountryAliasTable)
            {
                string[] parts = entry.Split('|');
                if (!countryNameToCode.ContainsKey(parts[1]))
                    countryNameToCode[parts[1]] = parts[0];
                if (!countryCodeToName.ContainsKey(parts[0]))
                    countryCodeToName[parts[0]] = parts[1];
            }

            _stateNameToCode = stateNameToCode;
            _stateCodeToName = stateCodeToName;
            _countryNameToCode = countryNameToCode;
            _countryCodeToName = countryCodeToName;
        }

        private string FormatRegion(string value, string format)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;

            EnsureRegionMaps();

            string text = value.Trim().TrimEnd('.');
            string result;

            switch (format)
            {
                case "State name to code":
                    return _stateNameToCode.TryGetValue(text, out result) ? result : value;

                case "State code to name":
                    return _stateCodeToName.TryGetValue(text, out result) ? result : value;

                case "Country name to code":
                    return _countryNameToCode.TryGetValue(text, out result) ? result : value;

                case "Country code to name":
                    return _countryCodeToName.TryGetValue(text, out result) ? result : value;
            }

            return value;
        }

        // Date parsing, Input patterns for a locale
        private static string[] GetInputFormats(string locale)
        {
            return string.Equals(locale, LocaleGlobal, StringComparison.OrdinalIgnoreCase)
                ? GlobalInputFormats
                : UsInputFormats;
        }

        // Single parse routine used by both the text and the serial conversions
        private static bool TryParseDate(object value, string locale, out DateTime result)
        {
            result = DateTime.MinValue;

            if (value == null) return false;

            if (value is DateTime dateTimeValue)
            {
                result = dateTimeValue;
                return true;
            }

            // Excel serial date
            if (value is double numericValue)
            {
                try
                {
                    result = DateTime.FromOADate(numericValue);
                    return true;
                }
                catch (ArgumentException)
                {
                    return false;
                }
            }

            string text = value.ToString().Trim();
            if (text.Length == 0) return false;

            if (DateTime.TryParseExact(text, GetInputFormats(locale), CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces, out result))
            {
                return true;
            }

            // Last resort: let the matching culture take a pass at anything the
            // explicit list missed, e.g. "Aug 27 2026" or "27 August 2026"
            CultureInfo culture = string.Equals(locale, LocaleGlobal, StringComparison.OrdinalIgnoreCase)
                ? CultureInfo.GetCultureInfo("en-GB")
                : CultureInfo.GetCultureInfo("en-US");

            return DateTime.TryParse(text, culture, DateTimeStyles.AllowWhiteSpaces, out result);
        }

        // Formats that ToString cannot produce, so they are built by hand
        private static bool IsSpecialDateFormat(string format)
        {
            return !string.IsNullOrEmpty(format) && SpecialDateFormats.Contains(format);
        }

        // ISO 8601 week. System.Globalization.ISOWeek is .NET Core only,
        // so the Thursday rule is applied directly
        private static DateTime IsoThursday(DateTime date)
        {
            int offset = 3 - (((int)date.DayOfWeek + 6) % 7);
            return date.AddDays(offset);
        }

        private static int IsoWeekNumber(DateTime date)
        {
            DateTime thursday = IsoThursday(date);
            return ((thursday.DayOfYear - 1) / 7) + 1;
        }

        private static int IsoWeekYear(DateTime date)
        {
            return IsoThursday(date).Year;
        }

        private static string FormatSpecialDate(DateTime date, string format)
        {
            int quarter = ((date.Month - 1) / 3) + 1;

            switch (format)
            {
                case FmtQuarterFirst:
                    return $"Q{quarter} {date.Year}";

                case FmtQuarterLast:
                    return $"{date.Year}-Q{quarter}";

                case FmtIsoWeek:
                    return $"{IsoWeekYear(date)}-W{IsoWeekNumber(date):00}";

                case FmtOrdinalDay:
                    return $"{date.Year}-{date.DayOfYear:000}";
            }

            return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        // A one-character custom format is read by .NET as a STANDARD specifier
        // ("d" = short date), so escape it to keep the custom meaning
        private static string SafeOutputFormat(string dateFormat)
        {
            if (string.IsNullOrEmpty(dateFormat)) return "yyyy-MM-dd";
            return dateFormat.Length == 1 ? "%" + dateFormat : dateFormat;
        }

        // Method to format date based on a specified format
        private string FormatDateToString(object value, string dateFormat, string locale)
        {
            if (value == null) return string.Empty;

            if (TryParseDate(value, locale, out DateTime parsedDate))
            {
                if (IsSpecialDateFormat(dateFormat))
                {
                    return FormatSpecialDate(parsedDate, dateFormat);
                }

                return parsedDate.ToString(SafeOutputFormat(dateFormat), CultureInfo.InvariantCulture);
            }

            return value.ToString(); // Return original if not a valid date
        }

        // Method to convert an object to the Excel date serial
        private object ConvertToExcelSerial(object value, string locale)
        {
            // If value is already a double (assumed to be an Excel serial date), return it
            if (value is double numericValue)
            {
                return numericValue;
            }

            if (TryParseDate(value, locale, out DateTime parsedDate))
            {
                return parsedDate.ToOADate();
            }

            // If the value cannot be converted, return the original value
            return value;
        }

        // .NET format string -> Excel number format code
        private static string ToExcelNumberFormat(string netFormat)
        {
            if (string.IsNullOrWhiteSpace(netFormat)) return "General";

            // Quarters, fiscal years and week numbers stay as text
            if (IsSpecialDateFormat(netFormat)) return "@";

            var sb = new StringBuilder();
            bool has12Hour = false;
            bool hasMeridiem = false;
            int i = 0;

            while (i < netFormat.Length)
            {
                char c = netFormat[i];

                // Escaped character: \d
                if (c == '\\' && i + 1 < netFormat.Length)
                {
                    sb.Append('"').Append(netFormat[i + 1]).Append('"');
                    i += 2;
                    continue;
                }

                // Quoted literal: 'of' or "of"
                if (c == '\'' || c == '"')
                {
                    char quote = c;
                    i++;
                    var literal = new StringBuilder();
                    while (i < netFormat.Length && netFormat[i] != quote)
                    {
                        literal.Append(netFormat[i]);
                        i++;
                    }
                    if (i < netFormat.Length) i++; // closing quote
                    if (literal.Length > 0) sb.Append('"').Append(literal).Append('"');
                    continue;
                }

                // Consume the whole run of this character
                int runLength = 1;
                while (i + runLength < netFormat.Length && netFormat[i + runLength] == c) runLength++;
                string run = new string(c, runLength);

                switch (c)
                {
                    case 'd':                                          // d dd ddd dddd
                        sb.Append(run);
                        break;

                    case 'M':                                          // month -> lowercase
                        sb.Append(new string('m', runLength));
                        break;

                    case 'y':                                          // Excel has only yy / yyyy
                        sb.Append(runLength <= 2 ? "yy" : "yyyy");
                        break;

                    case 'H':                                          // 24-hour
                        sb.Append(new string('h', runLength));
                        break;

                    case 'h':                                          // 12-hour, needs AM/PM
                        has12Hour = true;
                        sb.Append(run);
                        break;

                    case 'm':                                          // minutes
                    case 's':
                        sb.Append(run);
                        break;

                    case 'f':                                          // fractional seconds, Excel caps at 3
                    case 'F':
                        sb.Append('.').Append(new string('0', Math.Min(runLength, 3)));
                        break;

                    case 't':
                        hasMeridiem = true;
                        sb.Append("AM/PM");
                        break;

                    case ':':
                    case '/':
                    case '-':
                    case ' ':
                        sb.Append(run);
                        break;

                    case '.':                                          // quote: bare "." is a decimal point
                    case ',':                                          // quote: bare "," is a thousands separator
                        sb.Append('"').Append(run).Append('"');
                        break;

                    default:                                           // literal text, e.g. the T in ISO 8601
                        sb.Append('"').Append(run).Append('"');
                        break;
                }

                i += runLength;
            }

            string excelFormat = sb.ToString();

            // 12-hour clock without a designator renders 13:00 as 1:00 in Excel
            if (has12Hour && !hasMeridiem)
            {
                excelFormat += " AM/PM";
            }

            return string.IsNullOrWhiteSpace(excelFormat) ? "General" : excelFormat;
        }

        // Digit mask -> Excel custom number format, e.g. (###) ###-#### -> (000) 000-0000
        // Keeps leading zeros on IDs that are stored as real numbers
        private static string ToExcelDigitMask(string format)
        {
            if (string.IsNullOrWhiteSpace(format)) return "General";
            if (format.IndexOf('#') < 0) return "@";

            return format.Replace('#', '0');
        }

        // True when the category holds digits that Excel can carry as a number
        private static bool IsDigitCategory(string category)
        {
            return category == CatPhone || category == CatZip
                || category == CatSsn || category == CatEin;
        }

        // Excel reads / as a fraction separator, . as a decimal point and , as a
        // thousands separator, and a bare digit in a format code is a literal
        // rather than a placeholder. Only masks built from these characters
        // survive the trip through NumberFormat, so the rest are text only
        private static bool IsExcelSafeDigitMask(string format)
        {
            if (string.IsNullOrEmpty(format) || format.IndexOf('#') < 0) return false;

            foreach (char c in format)
            {
                if (c != '#' && c != ' ' && c != '(' && c != ')' && c != '-')
                {
                    return false;
                }
            }

            return true;
        }

        // Whether "Excel Format" is a valid choice for this category and format
        private static bool SupportsExcelFormat(string category, string format)
        {
            if (string.IsNullOrEmpty(category) || string.IsNullOrEmpty(format)) return false;

            // Quarters, week numbers and ordinal days have no serial equivalent
            if (category == CatDate) return !IsSpecialDateFormat(format);

            if (IsDigitCategory(category)) return IsExcelSafeDigitMask(format);

            // Email, durations and region names have no numeric form at all
            return false;
        }

        // Main conversion method for all categories
        private void ConvertValues(Excel.Application excelApp, string category, string format,
            string convertType, string convertLocale, string currentLocale)
        {
            Excel.Range rangeToProcess = QTUtils.GetRangeToProcess(excelApp);
            if (rangeToProcess == null) return;

            try
            {
                excelApp.ScreenUpdating = false;

                // Only reformat what the user can actually see
                Excel.Range formatTarget = VisiblePortion(rangeToProcess);
                SetInitialNumberFormat(formatTarget, convertType);

                // Filter-aware: values are compacted, sourceRows maps each back to its sheet row
                var (values, sourceRows) = QTUtils.GetRangeValuesWithRows(rangeToProcess);

                if (ProcessValues(values, category, format, convertType, currentLocale))
                {
                    SetFinalNumberFormat(formatTarget, convertType, category, format);
                    QTUtils.SetRangeValues(rangeToProcess, values, sourceRows);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing values: {ex.Message}");
                Debug.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
            finally
            {
                excelApp.ScreenUpdating = true;
            }
        }

        // Visible portion of a range when a filter is hiding rows, else the range itself
        private Excel.Range VisiblePortion(Excel.Range range)
        {
            try
            {
                if (range.Worksheet.AutoFilterMode)
                    return range.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
            }
            catch { /* SpecialCells throws when nothing matches */ }

            return range;
        }

        private bool ProcessValues(object[,] values, string category, string format,
            string convertType, string currentLocale)
        {
            if (values == null || values.Length == 0) return false;

            bool modified = false;

            // Bounds per dimension - GetRangeValuesWithRows returns 0-based arrays
            // while Value2 returns 1-based, so never reuse dimension 0's bound for columns
            int firstRow = values.GetLowerBound(0);
            int lastRow = values.GetUpperBound(0);
            int firstCol = values.GetLowerBound(1);
            int lastCol = values.GetUpperBound(1);

            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int col = firstCol; col <= lastCol; col++)
                {
                    if (TryProcessCell(values, row, col, category, format, convertType, currentLocale))
                    {
                        modified = true;
                    }
                }
            }

            return modified;
        }

        private void SetInitialNumberFormat(Excel.Range range, string convertType)
        {
            range.NumberFormat = convertType == "Text" ? "@" : "General";
        }

        private bool TryProcessCell(object[,] values, int row, int col, string category,
            string format, string convertType, string currentLocale)
        {
            var cellValue = values[row, col];
            if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                return false;

            values[row, col] = convertType == "Text"
                ? FormatAsText(cellValue, category, format, currentLocale)
                : FormatAsExcel(cellValue, category, format, currentLocale);

            return true;
        }

        private object FormatAsText(object value, string category, string format, string currentLocale)
        {
            switch (category)
            {
                case CatDate:
                    return FormatDateToString(value, format, currentLocale);

                case CatPhone:
                    return FormatPhone(value.ToString(), format);

                case CatZip:
                    return FormatZip(value.ToString(), format);

                case CatSsn:
                    return FormatSSN(value.ToString(), format);

                case CatEin:
                    return FormatEIN(value.ToString(), format);

                case CatEmail:
                    return FormatEmail(value.ToString(), format);

                case CatTime:
                    // Passed as an object so Excel's time fractions survive
                    return FormatTime(value, format);

                case CatRegion:
                    return FormatRegion(value.ToString(), format);
            }

            return value;
        }

        private object FormatAsExcel(object value, string category, string format, string currentLocale)
        {
            if (category == CatDate)
            {
                // Quarters and week numbers have no serial equivalent
                if (IsSpecialDateFormat(format))
                {
                    return FormatDateToString(value, format, currentLocale);
                }

                return ConvertToExcelSerial(value, currentLocale);
            }

            if (IsDigitCategory(category) && SupportsExcelFormat(category, format))
            {
                return DigitsOnly(value.ToString());
            }

            // Everything else has no meaningful numeric form
            return FormatAsText(value, category, format, currentLocale);
        }

        private void SetFinalNumberFormat(Excel.Range range, string convertType,
            string category, string format)
        {
            if (convertType != "Excel Format") return;

            if (category == CatDate)
            {
                range.NumberFormat = ToExcelNumberFormat(format);
                return;
            }

            if (IsDigitCategory(category) && SupportsExcelFormat(category, format))
            {
                range.NumberFormat = ToExcelDigitMask(format);
                return;
            }

            range.NumberFormat = "@";
        }

        // Only offer "Excel Format" where the selected format can actually
        // survive as an Excel number format, otherwise leave Text as the only choice
        private void UpdateConvertTypeOptions()
        {
            string category = CbCategory.SelectedItem?.ToString();
            string format = CbFormat.SelectedItem?.ToString();
            bool allowExcel = SupportsExcelFormat(category, format);

            string previous = CbConvertType.SelectedItem?.ToString() ?? "Text";

            CbConvertType.Items.Clear();
            CbConvertType.Items.Add("Text");
            if (allowExcel) CbConvertType.Items.Add("Excel Format");

            CbConvertType.SelectedItem = allowExcel && previous == "Excel Format"
                ? "Excel Format"
                : "Text";

            CbConvertType.Enabled = allowExcel;
        }

        // Sample value that matches the selected format. Categories whose
        // formats convert in opposite directions need a different sample per format
        private static string DefaultExample(string category, string format)
        {
            if (string.IsNullOrEmpty(format)) return null;

            if (category == CatRegion)
            {
                switch (format)
                {
                    case "State name to code": return "Massachusetts";
                    case "State code to name": return "MA";
                    case "Country name to code": return "United States";
                    case "Country code to name": return "US";
                }

                return null;
            }

            if (category == CatZip)
            {
                switch (format)
                {
                    case "#####": return "02215-9997";      // shows the ZIP+4 being trimmed
                    case "#####-####": return "022159997";
                    case "#########": return "02215-9997";
                    case "A1A 1A1": return "K1A0B1";
                    case "AA9A 9AA": return "SW1A1AA";
                }

                return null;
            }

            return null;
        }

        // Example preview. Method to update Example Textbox
        private void UpdateExample(object sender, EventArgs e)
        {
            string category = CbCategory.SelectedItem?.ToString();
            string format = CbFormat.SelectedItem?.ToString();

            if (string.IsNullOrEmpty(category) || string.IsNullOrEmpty(format)
                || string.IsNullOrWhiteSpace(TbExample.Text))
            {
                return;
            }

            // Falls back to US rather than throwing when the locale box is blanked
            string currentLocale = CbCurrentLocale.SelectedItem?.ToString() ?? LocaleUs;

            TbExFormatted.Text = FormatAsText(TbExample.Text, category, format, currentLocale)?.ToString()
                ?? string.Empty;
        }

        // Example change event
        private void TbExample_TextChanged(object sender, EventArgs e)
        {
            UpdateExample(sender, e);
        }

        // Format change event
        private void CbFormat_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateConvertTypeOptions();

            string category = CbCategory.SelectedItem?.ToString();
            string format = CbFormat.SelectedItem?.ToString();

            string example = DefaultExample(category, format);
            if (example != null && !string.Equals(TbExample.Text, example, StringComparison.Ordinal))
            {
                // Assigning the text raises TbExample_TextChanged, which refreshes the preview
                TbExample.Text = example;
                return;
            }

            UpdateExample(sender, e);
        }

        // Convert locale change event
        private void CbConvertLocale_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFieldOptions();
        }

        // Current locale change event
        private void CbCurrentLocale_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFieldOptions();
        }

    }
}