using System;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    public partial class TextConvertForm : Form
    {
        private readonly Excel.Application _excelApp;
   
        public TextConvertForm(Excel.Application excelApp)
        {
            InitializeComponent();
            _excelApp = excelApp;
        }

        // On Load
        private void TextConvertForm_Load(object sender, EventArgs e)
        {
            // Populate the converstion type
            this.CbConvertType.Items.AddRange(new object[]
            {
                "Text",
                "Excel Format"
            });

            // Set the default converstion type
            this.CbConvertType.SelectedItem = "Text";

            // Populate the current locale
            this.CbCurrentLocale.Items.AddRange(new object[]
            {
                "USA (Month, Day, Year)",
                "Other (Day, Month, Year)"
            });

            // Set the default locale
            this.CbCurrentLocale.SelectedItem = "USA (Month, Day, Year)";

            // Populate the convert locale
            this.CbConvertLocale.Items.AddRange(new object[]
            {
                "USA (Month, Day, Year)",
                "Other (Day, Month, Year)"
            });

            // Set the default locale
            this.CbConvertLocale.SelectedItem = "USA (Month, Day, Year)";

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
            string format = CbFormat.SelectedItem?.ToString() ?? string.Empty;

            // Ensure category is not null or empty before proceeding
            if (string.IsNullOrEmpty(category)) return;

            // Populate CbFormat based on the selected item in CbCategory
            switch (category)
            {
                case "Date":
                    if (convertLocale.Equals("Other (Day, Month, Year)"))
                    {
                        // Populate DMY date formats
                        this.CbFormat.Items.AddRange(new object[]
                        {
                            "yyyy-MM-dd",
                            "dd/MM/yyyy",
                            "d/M/yyyy",
                            "dd MMM yyyy",
                            "dd MMMM yyyy",
                            "yyyy/MM/dd",
                            "yyyy.MM.dd",
                            "yyyy MMM dd"
                        });
                    } 
                    else 
                    { 
                        this.CbFormat.Items.AddRange(new object[]
                        {
                            "yyyy-MM-dd",
                            "M/d/yyyy",
                            "MM/dd/yyyy",
                            "MMM dd, yyyy",
                            "MMMM dd, yyyy",
                            "yyyy/MM/dd",
                            "yyyy.MM.dd",
                            "yyyy MMM dd"
                        });
                    }
                    if(currentLocale.Equals("Other (Day, Month, Year)"))
                    {
                        // DMY Example
                        this.TbExample.Text = DateTime.Now.ToString("d/M/yyyy");
                    }
                    else
                    {
                        this.TbExample.Text = DateTime.Now.ToString("M/d/yyyy");
                    }
                    this.CbCurrentLocale.Enabled = true;
                    this.CbConvertLocale.Enabled = true;
                    this.LblCurrentLocale.Enabled = true;
                    this.LblFormatLocale.Enabled = true;

                    // If null from category change then set it
                    if (string.IsNullOrEmpty(this.CbCurrentLocale.Text))
                    {
                        this.CbCurrentLocale.Text = "USA (Month, Day, Year)";
                    }
                    if (string.IsNullOrEmpty(this.CbConvertLocale.Text))
                    {
                        this.CbConvertLocale.Text = "USA (Month, Day, Year)";
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
            Excel.Worksheet ws = _excelApp.ActiveSheet;
            Excel.Range selectedRange = _excelApp.Selection;

            // Create an instance of QTClipboard
            QTClipboard clipboard = QTClipboard.Instance;
            // Copy and store values
            clipboard.CopyAndStoreFormat(selectedRange);
            // Enable the Undo button
            Globals.Ribbons.Ribbon1.BtnUndo.Enabled = true;

            // Form values
            string category = CbCategory.SelectedItem.ToString();
            string format = CbFormat.SelectedItem.ToString();
            string convertType = CbConvertType.SelectedItem.ToString();
            string currentLocale = CbCurrentLocale.SelectedItem.ToString();
            string convertLocale = CbConvertLocale.SelectedItem.ToString();

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

        // Method to format date based on a specified format
        private string FormatDateToString(object value, string dateFormat, string locale)
        {
            string[] inputDateFormats;

            // Define date formats based on locale
            if (locale.Equals("USA (Month, Day, Year)"))
            {
                inputDateFormats = new[]
                {
            "yyyy-MM-dd", "M/d/yyyy", "MM/dd/yyyy", "MMM dd, yyyy", "MMMM dd, yyyy",
            "yyyy/MM/dd", "yyyy.MM.dd", "yyyy MMM dd", "M/d/yyyy H:mm", "MM/dd/yyyy H:mm",
            "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss.ffffffZ", "MMM dd, yyyy HH:mm",
            "MMMM dd, yyyy h:mm"
        };
            }
            else
            {
                inputDateFormats = new[]
                {
            "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd MMM yyyy", "dd MMMM yyyy",
            "yyyy/MM/dd", "yyyy.MM.dd", "yyyy MMM dd", "d/M/yyyy H:mm", "dd/MM/yyyy H:mm",
            "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss.ffffffZ", "dd MMM yyyy HH:mm",
            "dd MMMM yyyy h:mm"
        };
            }

            if (value == null)
                return string.Empty;

            // Check if the value is a valid date string
            if (value is string dateString && DateTime.TryParseExact(dateString, inputDateFormats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
            {
                return parsedDate.ToString(dateFormat);
            }
            // Handle numeric Excel date serial numbers
            else if (value is double numericValue)
            {
                DateTime dateValue = DateTime.FromOADate(numericValue);
                return dateValue.ToString(dateFormat);
            }

            return value.ToString(); // Return original if not a valid date
        }

        // Method to convert an object to the Excel date serial
        private object ConvertToExcelSerial(object value, string locale)
        {
            string[] inputDateFormats;

            // Define date formats based on locale
            if (locale.Equals("USA (Month, Day, Year)"))
            {
                inputDateFormats = new[]
                {
            "yyyy-MM-dd", "M/d/yyyy", "MM/dd/yyyy", "MMM dd, yyyy", "MMMM dd, yyyy",
            "yyyy/MM/dd", "yyyy.MM.dd", "yyyy MMM dd", "M/d/yyyy H:mm", "MM/dd/yyyy H:mm",
            "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss.ffffffZ", "MMM dd, yyyy HH:mm",
            "MMMM dd, yyyy h:mm"
        };
            }
            else
            {
                inputDateFormats = new[]
                {
            "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd MMM yyyy", "dd MMMM yyyy",
            "yyyy/MM/dd", "yyyy.MM.dd", "yyyy MMM dd", "d/M/yyyy H:mm", "dd/MM/yyyy H:mm",
            "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss.ffffffZ", "dd MMM yyyy HH:mm",
            "dd MMMM yyyy h:mm"
        };
            }

            // If value is already a double (assumed to be an Excel serial date), return it
            if (value is double numericValue)
            {
                return numericValue;
            }
            // If value is a string, attempt to parse it as a date
            else if (value is string dateString &&
                     DateTime.TryParseExact(
                         dateString,
                         inputDateFormats,
                         System.Globalization.CultureInfo.InvariantCulture,
                         System.Globalization.DateTimeStyles.None,
                         out DateTime parsedDate))
            {
                // Convert the parsed date to an Excel serial date
                return parsedDate.ToOADate();
            }

            // If the value cannot be converted, return the original value
            return value;
        }

        // Main conversion method for all categories
        private void ConvertValues(Excel.Application excelApp, string category, string format,
            string convertType, string convertLocale, string currentLocale)
        {
            Excel.Worksheet ws = excelApp.ActiveSheet;
            Excel.Range rangeToProcess = QTUtils.GetRangeToProcess(excelApp);

            try
            {
                excelApp.ScreenUpdating = false;
                SetInitialNumberFormat(rangeToProcess, convertType);

                var values = GetRangeValues(rangeToProcess);
                if (ProcessValues(values, category, format, convertType, currentLocale))
                {
                    SetFinalNumberFormat(rangeToProcess, convertType, format);
                    rangeToProcess.Value2 = values;
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

        private void SetInitialNumberFormat(Excel.Range range, string convertType)
        {
            range.NumberFormat = convertType == "Text" ? "@" : "General";
        }

        private object[,] GetRangeValues(Excel.Range range)
        {
            if (range.Rows.Count == 1 && range.Columns.Count == 1)
            {
                var values = new object[1, 1];
                values[0, 0] = range.Value2;
                return values;
            }
            return range.Value2 as object[,];
        }

        private bool ProcessValues(object[,] values, string category, string format,
            string convertType, string currentLocale)
        {
            if (values == null) return false;

            bool modified = false;
            int baseIndex = values.GetLowerBound(0);
            int rowsCount = values.GetUpperBound(0);
            int colsCount = values.GetUpperBound(1);

            for (int row = baseIndex; row <= rowsCount; row++)
            {
                for (int col = baseIndex; col <= colsCount; col++)
                {
                    if (TryProcessCell(values, row, col, category, format, convertType, currentLocale))
                    {
                        modified = true;
                    }
                }
            }
            return modified;
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
                range.NumberFormat = QTUtils.GetExcelNumberFormat(format);
            }
        }

        // Method to update Example Textbox
        private void UpdateExample(object sender, EventArgs e)
        {
            if (CbFormat.SelectedItem != null && !string.IsNullOrWhiteSpace(TbExample.Text))
            {
                switch (CbCategory.SelectedItem.ToString())
                {
                    case "Date":
                        TbExFormatted.Text = FormatDateToString(TbExample.Text.ToString(), CbFormat.SelectedItem.ToString(), CbCurrentLocale.SelectedItem.ToString());
                        break;

                    case "Phone Number":
                        TbExFormatted.Text = FormatPhone(TbExample.Text.ToString(), CbFormat.SelectedItem.ToString());
                        break;

                    case "Zip Code":
                        TbExFormatted.Text = FormatZip(TbExample.Text.ToString(), CbFormat.SelectedItem.ToString());
                        break;

                    case "Social Security Number":
                        TbExFormatted.Text = FormatSSN(TbExample.Text.ToString(), CbFormat.SelectedItem.ToString());
                        break;
                }

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
