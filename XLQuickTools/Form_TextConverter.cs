using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    public partial class TextConvertForm : Form
    {
        private readonly Excel.Application _excelApp;

        private const string LocaleUs = "United States";
        private const string LocaleGlobal = "Global";

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
            "MM/dd/yyyy h:mm tt"
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
            "dd/MM/yyyy HH:mm"
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
                "Date",
                "Phone Number",
                "Zip Code",
                "Social Security Number"
            });

            // Set the default category
            this.CbCategory.SelectedItem = "Date";

            // Subscribe to the SelectedIndexChanged event
            this.CbCategory.SelectedIndexChanged += CbCategory_SelectedIndexChanged;
        }

        // Category changed
        private void CbCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFieldOptions();
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
                case "Date":
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

                    this.CbCurrentLocale.Enabled = true;
                    this.CbConvertLocale.Enabled = true;
                    this.LblCurrentLocale.Enabled = true;
                    this.LblFormatLocale.Enabled = true;

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

                case "Phone Number":
                    this.CbFormat.Items.AddRange(new object[]
                    {
                        "(###) ###-####",
                        "###/###-####",
                        "###-###-####",
                        "+1 (###) ###-####"
                    });
                    this.TbExample.Text = "1234567890";
                    this.CbConvertType.SelectedItem = "Text";
                    this.CbCurrentLocale.Text = "";
                    this.CbConvertLocale.Text = "";
                    this.CbCurrentLocale.Enabled = false;
                    this.CbConvertLocale.Enabled = false;
                    this.LblCurrentLocale.Enabled = false;
                    this.LblFormatLocale.Enabled = false;
                    break;

                case "Zip Code":
                    this.CbFormat.Items.AddRange(new object[]
                    {
                        "#####-####"
                    });
                    this.TbExample.Text = "123451234";
                    this.CbConvertType.SelectedItem = "Text";
                    this.CbCurrentLocale.Text = "";
                    this.CbConvertLocale.Text = "";
                    this.CbCurrentLocale.Enabled = false;
                    this.CbConvertLocale.Enabled = false;
                    this.LblCurrentLocale.Enabled = false;
                    this.LblFormatLocale.Enabled = false;
                    break;

                case "Social Security Number":
                    this.CbFormat.Items.AddRange(new object[]
                    {
                        "###-##-####"
                    });
                    this.TbExample.Text = "123456789";
                    this.CbConvertType.SelectedItem = "Text";
                    this.CbCurrentLocale.Text = "";
                    this.CbConvertLocale.Text = "";
                    this.CbCurrentLocale.Enabled = false;
                    this.CbConvertLocale.Enabled = false;
                    this.LblCurrentLocale.Enabled = false;
                    this.LblFormatLocale.Enabled = false;
                    break;
            }

            // Set a default format
            if (CbFormat.Items.Count > 0)
            {
                CbFormat.SelectedIndex = 0;
            }
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

        // Method to format a phone number based on the specified format
        private string FormatPhone(string phoneNumber, string format)
        {
            if (string.IsNullOrWhiteSpace(phoneNumber))
            {
                return string.Empty;
            }

            // Extract digits from the phone number
            string digitsOnly = DigitsOnly(phoneNumber);

            // Format the phone number based on format
            if (digitsOnly.Length == 10 || (digitsOnly.Length == 11 && format.StartsWith("+1")))
            {
                switch (format)
                {
                    case "###/###-####":
                        if (digitsOnly.Length == 10)
                        {
                            return $"{digitsOnly.Substring(0, 3)}/{digitsOnly.Substring(3, 3)}-{digitsOnly.Substring(6, 4)}";
                        }
                        break;

                    case "(###) ###-####":
                        if (digitsOnly.Length == 10)
                        {
                            return $"({digitsOnly.Substring(0, 3)}) {digitsOnly.Substring(3, 3)}-{digitsOnly.Substring(6, 4)}";
                        }
                        break;

                    case "###-###-####":
                        if (digitsOnly.Length == 10)
                        {
                            return $"{digitsOnly.Substring(0, 3)}-{digitsOnly.Substring(3, 3)}-{digitsOnly.Substring(6, 4)}";
                        }
                        break;

                    case "+1 (###) ###-####":
                        if (digitsOnly.Length == 10)
                        {
                            return $"+1 ({digitsOnly.Substring(0, 3)}) {digitsOnly.Substring(3, 3)}-{digitsOnly.Substring(6, 4)}";
                        }
                        else if (digitsOnly.Length == 11 && digitsOnly.StartsWith("1"))
                        {
                            return $"+1 ({digitsOnly.Substring(1, 3)}) {digitsOnly.Substring(4, 3)}-{digitsOnly.Substring(7, 4)}";
                        }
                        break;
                }
            }

            // If the format doesn't match, return the original phone number
            return phoneNumber;
        }

        // Method to format ZIP code
        private string FormatZip(string zipCode, string format)
        {
            if (string.IsNullOrEmpty(zipCode)) return string.Empty;

            // Extract digits only
            string digitsOnly = DigitsOnly(zipCode);

            // Format ZIP code based on length
            if (digitsOnly.Length == 4)
            {
                // Pad with 0 for 5-digit ZIP code
                return $"0{digitsOnly}";
            }
            else if (digitsOnly.Length == 8)
            {
                // Pad with 0 and format as ZIP+4
                return $"0{digitsOnly.Substring(0, 4)}-{digitsOnly.Substring(4, 4)}";
            }
            else if (digitsOnly.Length == 9)
            {
                // ZIP+4 format
                return $"{digitsOnly.Substring(0, 5)}-{digitsOnly.Substring(5, 4)}";
            }

            return zipCode; // Return original if no valid format
        }

        // Method to format a single SSN string
        private string FormatSSN(string ssn, string format)
        {
            if (string.IsNullOrEmpty(ssn)) return string.Empty;

            // Extract digits only
            string digitsOnly = DigitsOnly(ssn);

            // Format SSN
            if (digitsOnly.Length == 9)
            {
                return $"{digitsOnly.Substring(0, 3)}-{digitsOnly.Substring(3, 2)}-{digitsOnly.Substring(5, 4)}";
            }
            else if (digitsOnly.Length == 8)
            {
                return $"0{digitsOnly.Substring(0, 2)}-{digitsOnly.Substring(2, 2)}-{digitsOnly.Substring(4, 4)}";
            }

            return ssn; // Return original if no valid format
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
                    SetFinalNumberFormat(formatTarget, convertType, format);
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
                : FormatAsExcel(cellValue, category, currentLocale);

            return true;
        }

        private object FormatAsText(object value, string category, string format, string currentLocale)
        {
            if (category == "Date")
            {
                return FormatDateToString(value, format, currentLocale);
            }
            else if (category == "Phone Number")
            {
                return FormatPhone(value.ToString(), format);
            }
            else if (category == "Zip Code")
            {
                return FormatZip(value.ToString(), format);
            }
            else if (category == "Social Security Number")
            {
                return FormatSSN(value.ToString(), format);
            }
            return value;
        }

        private object FormatAsExcel(object value, string category, string currentLocale)
        {
            if (category == "Date")
            {
                return ConvertToExcelSerial(value, currentLocale);
            }
            return DigitsOnly(value.ToString());
        }

        private void SetFinalNumberFormat(Excel.Range range, string convertType, string format)
        {
            if (convertType == "Excel Format")
            {
                range.NumberFormat = ToExcelNumberFormat(format);
            }
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

            switch (category)
            {
                case "Date":
                    TbExFormatted.Text = FormatDateToString(TbExample.Text, format, currentLocale);
                    break;

                case "Phone Number":
                    TbExFormatted.Text = FormatPhone(TbExample.Text, format);
                    break;

                case "Zip Code":
                    TbExFormatted.Text = FormatZip(TbExample.Text, format);
                    break;

                case "Social Security Number":
                    TbExFormatted.Text = FormatSSN(TbExample.Text, format);
                    break;
            }
        }

        // Example change event
        private void TbExample_TextChanged(object sender, EventArgs e)
        {
            UpdateExample(sender, e);
        }

        // Format change event
        private void CbFormat_SelectedIndexChanged(object sender, EventArgs e)
        {
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