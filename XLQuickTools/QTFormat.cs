using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using static XLQuickTools.QTSettings;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Text;

namespace XLQuickTools
{
    public class QTFormat
    {
        // Method to transform text based on an option
        private static string TransformText(string input, int option, string leading = null, string trailing = null)
        {
            switch (option)
            {
                case 0: // Uppercase
                    return input.ToUpper();
                case 1: // Lowercase
                    return input.ToLower();
                case 2: // Proper case
                    return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(input.ToLower());
                case 3: // Remove letters
                    return System.Text.RegularExpressions.Regex.Replace(input, "\\p{L}", "");
                case 4: // Remove numbers
                    return System.Text.RegularExpressions.Regex.Replace(input, "[0-9]", "");
                case 5: // Remove special characters (keep spaces and accents)
                    return System.Text.RegularExpressions.Regex.Replace(input, @"[^a-zA-Z0-9\u00C0-\u024F\s]", "");
                case 6: // Normalize Text/Remove diacritics
                    return NormalizeText(input);
                case 7: // Trim and Clean
                    if (input != null)
                    {
                        // Trim
                        input = input.Trim();

                        // Check for string containing only spaces
                        if (input.Length > 0 && input.All(c => c == ' '))
                        {
                            input = string.Empty;
                        }
                    }
                    return Clean(input);
                case 8: // Add leading or trailing
                    return (leading ?? "") + input + (trailing ?? "");
                case 9: // Remove non-ASCII
                    return ReplaceNonAscii(input);
                default:
                    throw new ArgumentException("Invalid option.");
            }
        }

        // Format select menu
        public static void FormatMenu(int option, string leading = null, string trailing = null)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range rangeToProcess = null;

            try
            {
                rangeToProcess = QTUtils.GetRangeToProcess(excelApp);
                if (rangeToProcess == null) return;

                excelApp.ScreenUpdating = false;
                var values = QTUtils.GetRangeValues(rangeToProcess);

                // Create an instance of QTClipboard
                QTClipboard clipboard = QTClipboard.Instance;
                // Copy and store values
                clipboard.CopyAndStoreFormat(rangeToProcess);

                if (ProcessFormat(values, option, leading, trailing))
                {
                    rangeToProcess.Value2 = values;
                }
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
            finally
            {
                excelApp.ScreenUpdating = true;
                QTUtils.CleanupResources(rangeToProcess);
            }
        }

        // Process the Format Menu option array and modify values if necessary
        private static bool ProcessFormat(object[,] values, int option, string leading = null, string trailing = null)
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
                    if (values[row, col] != null)
                    {
                        // Convert to string
                        string cellValue = values[row, col].ToString();
                        if (cellValue.Length > 0)
                        {
                            values[row, col] = TransformText(cellValue, option, leading, trailing).Trim();
                            modified = true;
                        }
                    }
                }
            }
            return modified;
        }

        // Method to remove excess formatting from entirie workbook
        public static void RemoveExcess()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;

            if (activeWorkbook != null)
            {
                try
                {
                    // Turn off screen updating
                    excelApp.ScreenUpdating = false;

                    foreach (Excel.Worksheet ws in activeWorkbook.Worksheets)
                    {
                        // Last row with actual data
                        int lastDataRow = ws.Cells.Find("*", System.Reflection.Missing.Value,
                            Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                            false, System.Reflection.Missing.Value, System.Reflection.Missing.Value)?.Row ?? 0;

                        // Last column with actual data
                        int lastDataColumn = ws.Cells.Find("*", System.Reflection.Missing.Value,
                            Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                            Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                            false, System.Reflection.Missing.Value, System.Reflection.Missing.Value)?.Column ?? 0;

                        // Clear excess rows
                        int totalRows = ws.Rows.Count;
                        if (lastDataRow < totalRows)
                        {
                            Excel.Range rowsToClear = ws.Rows[$"{lastDataRow + 1}:{totalRows}"];
                            rowsToClear.ClearFormats();
                        }

                        // Clear excess columns
                        int totalColumns = ws.Columns.Count;
                        if (lastDataColumn < totalColumns)
                        {
                            // Use a safe range
                            int startColumn = lastDataColumn + 1;
                            if (startColumn <= totalColumns)
                            {
                                Excel.Range columnsToClear = ws.Range[ws.Cells[1, startColumn], ws.Cells[1, totalColumns]];
                                columnsToClear.ClearFormats();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    QTUtils.ShowError(ex);
                }
                finally
                {
                    // Turn screen updating back on
                    excelApp.ScreenUpdating = true;
                    System.Windows.Forms.MessageBox.Show("Any excess formatting has been removed from the workbook.",
                        "Remove Excess", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
        }

        // Method for Trim & Clean (worksheet and workbook)
        public static void TrimClean(String rangeType)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            int option = 7; // Trim and Clean
            bool processSuccess = true;

            try
            {
                excelApp.ScreenUpdating = false;

                if (rangeType.ToLower() == "worksheet")
                {
                    // Process active worksheet only
                    Excel.Worksheet activeSheet = excelApp.ActiveSheet;
                    processSuccess = ProcessSheet(activeSheet, option);
                }
                else if (rangeType.ToLower() == "workbook")
                {
                    // Process all worksheets in workbook
                    foreach (Excel.Worksheet worksheet in activeWorkbook.Worksheets)
                    {
                        if (!ProcessSheet(worksheet, option))
                        {
                            processSuccess = false;
                            break;
                        }
                    }
                }
                else
                {
                    throw new ArgumentException("Invalid range type. Must be either 'worksheet' or 'workbook'.");
                }
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
                processSuccess = false;
            }
            finally
            {
                // Turn screen updating back on
                excelApp.ScreenUpdating = true;

                // Show appropriate message
                if (processSuccess)
                {
                    string scope = rangeType.ToLower() == "worksheet" ? "worksheet" : "workbook";
                    System.Windows.Forms.MessageBox.Show(
                        $"All cells in the {scope} have been trimmed and cleaned of any non-printable characters!",
                        "Trim and Clean",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
        }

        // Helper method to process individual worksheets
        private static bool ProcessSheet(Excel.Worksheet sheet, int option)
        {
            try
            {
                Excel.Range rangeToProcess = sheet.UsedRange;
                if (rangeToProcess == null) return true;

                var values = QTUtils.GetRangeValues(rangeToProcess);

                // Apply trim & clean
                if (ProcessFormat(values, option))
                {
                    rangeToProcess.Value2 = values;
                    return true;
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // Custom method to mimic Excel's CLEAN function
        public static string Clean(string input)
        {
            // Removes non-printable characters
            return new string(input.Where(c => !char.IsControl(c)).ToArray());
        }

        // Function to remove formatting
        public static void RemoveFormatting()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet ws = excelApp.ActiveSheet;

            try
            {
                // Apply to the entire sheet
                Excel.Range entireSheet = ws.Cells;

                // Clear all existing formatting in one call
                entireSheet.ClearFormats();

                // Retrieve and reset default font settings
                string defaultFontName = excelApp.StandardFont;
                double defaultFontSize = excelApp.StandardFontSize;
                entireSheet.Font.Name = defaultFontName;
                entireSheet.Font.Size = defaultFontSize;

                // Remove borders
                entireSheet.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                // Unfreeze panes and reset splits
                if (excelApp.ActiveWindow.FreezePanes)
                {
                    excelApp.ActiveWindow.FreezePanes = false;
                }
                excelApp.ActiveWindow.SplitRow = 0;
                excelApp.ActiveWindow.SplitColumn = 0;

                // Remove auto-filter if it exists
                if (ws.AutoFilterMode)
                {
                    ws.AutoFilterMode = false;
                }

                // Reset gridlines and zoom
                excelApp.ActiveWindow.DisplayGridlines = true;
                excelApp.ActiveWindow.Zoom = 100;

                // Reset row height and column width
                ws.Rows.RowHeight = ws.StandardHeight;
                ws.Columns.ColumnWidth = ws.StandardWidth;
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
        }

        // Remove formatting with screen updating turned off
        public static void RemoveFormattingNoUpdates()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            excelApp.ScreenUpdating = false;
            RemoveFormatting();
            excelApp.ScreenUpdating = true;
        }

        // Function to apply auto filter
        public static void ApplyFilter(Excel.Range usedRange)
        {
            // Check if the range contains more than one row and more than one column
            if (usedRange.Rows.Count > 1 && usedRange.Columns.Count > 1)
            {
                // Set the first row of the used range
                Excel.Range firstRow = usedRange.Rows[1];

                if (firstRow != null && firstRow.Cells.Count > 0)
                {
                    try
                    {
                        firstRow.AutoFilter(1); // Apply auto-filter to the first row
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        Debug.WriteLine("Autofilter cannot be applied: " + ex.Message);
                    }
                }
            }
        }

        // Apply Quick Format
        public static void QuickFormat()
        {
            // Get Excel application and active workbook
            var excelApp = Globals.ThisAddIn.Application;
            var activeWorkbook = excelApp.ActiveWorkbook;
            if (activeWorkbook == null) return;

            // Disable screen updating
            excelApp.ScreenUpdating = false;

            try
            {
                // Load settings and validate
                UserSettings settings = QTSettings.LoadUserSettingsFromXml();
                if (settings == null) return;

                // Get active sheet and used range
                var activeSheet = excelApp.ActiveSheet;
                var usedRange = QTUtils.GetEffectiveUsedRange(activeSheet);
                if (usedRange == null) return;

                // Remove existing formatting first
                RemoveFormatting();

                // Perform formatting operations
                ApplyRangeFormatting(usedRange, settings);
                ApplyWindowSettings(excelApp, settings);
                ApplyAlignmentSettings(usedRange, settings);
                // Apply these last (First row, Fit)
                ApplyFirstRowFormatting(usedRange, settings);
                ApplyFitFormatting(usedRange, settings);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Quick Format Error: {ex.Message}", "Excel Formatting Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Ensure screen updating is restored
                excelApp.ScreenUpdating = true;
            }
        }

        // Quick Format: First Row
        public static void ApplyFirstRowFormatting(Excel.Range usedRange, UserSettings settings)
        {
            if (usedRange == null || settings == null) return;

            var firstRow = usedRange.Rows[1];

            // Batch first row formatting operations
            if (settings.Bold || settings.Underline || settings.Border1 ||
                settings.InteriorColor || settings.AllignLeft_FR ||
                settings.AllignCenter_FR || settings.AllignRight_FR)
            {
                // Font formatting
                var firstRowFont = firstRow.Font;
                firstRowFont.Bold = settings.Bold;
                firstRowFont.Underline = settings.Underline;

                // Border
                if (settings.Border1)
                {
                    firstRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }

                // Interior Color
                if (settings.InteriorColor &&
                    int.TryParse(settings.Color, out int oleColor))
                {
                    firstRow.Interior.Color = ColorTranslator.ToOle(
                        ColorTranslator.FromOle(oleColor)
                    );
                }

                // Horizontal Alignment
                if (settings.AllignLeft_FR)
                    firstRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                else if (settings.AllignCenter_FR)
                    firstRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                else if (settings.AllignRight_FR)
                    firstRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                // Vertical Alignment
                if (settings.AllignTop_FR)
                    firstRow.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                else if (settings.AllignMiddle_FR)
                    firstRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                else if (settings.AllignBottom_FR)
                    firstRow.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
            }
        }

        public static void ApplyFitFormatting(Excel.Range usedRange, UserSettings settings)
        {
            // Formatting that needs to be completed at the end for things to fit properly

            // Apply AutoFilter if needed
            if (settings.AutoFilter)
            {
                ApplyFilter(usedRange);
            }

            // AutoFit Columns
            if (settings.AutoFit1)
            {
                usedRange.Columns.AutoFit();
            }

            // AutoFit Rows
            if (settings.AutoFit2)
            {
                usedRange.Rows.AutoFit();
            }
        }

        // Quick Format: Range Settings
        public static void ApplyRangeFormatting(Excel.Range usedRange, UserSettings settings)
        {
            if (usedRange == null || settings == null) return;

            // Batch range-wide formatting operations
            if (settings.Border2 || settings.WrapText ||
                settings.FontSize || settings.ColWidth ||
                settings.RowHeight || settings.AutoFit1 ||
                settings.AutoFit2)
            {
                // Borders
                if (settings.Border2)
                {
                    usedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }

                // Wrap Text
                usedRange.WrapText = settings.WrapText;

                // Font Size
                if (settings.FontSize &&
                    int.TryParse(settings.FontSizeTxt, out int fontSize))
                {
                    usedRange.Font.Size = fontSize;
                }

                // Column Width
                if (settings.ColWidth)
                {
                    usedRange.Columns.ColumnWidth = settings.Width;
                }

                // Row Height
                if (settings.RowHeight)
                {
                    usedRange.Rows.RowHeight = settings.Height;
                }
            }
        }

        // Quick Format: Window Settings
        public static void ApplyWindowSettings(Excel.Application excelApp, UserSettings settings)
        {
            if (excelApp == null || settings == null) return;

            // Freeze Panes
            if (settings.Freeze)
            {
                var activeWindow = excelApp.ActiveWindow;
                activeWindow.SplitRow = 1;
                activeWindow.FreezePanes = true;
            }

            // Zoom
            if (settings.Zoom &&
                int.TryParse(settings.ZoomValue, out int zoomValue))
            {
                excelApp.ActiveWindow.Zoom = zoomValue;
            }

            // Gridlines
            if (settings.Gridlines)
            {
                excelApp.ActiveWindow.DisplayGridlines = false;
            }
        }

        // Quick Format: Alignment Settings
        public static void ApplyAlignmentSettings(Excel.Range usedRange, UserSettings settings)
        {
            if (usedRange == null || settings == null) return;

            // Horizontal Alignment
            if (settings.AllignLeft)
                usedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            else if (settings.AllignCenter)
                usedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            else if (settings.AllignRight)
                usedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            // Vertical Alignment
            if (settings.AllignTop)
                usedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            else if (settings.AllignMiddle)
                usedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            else if (settings.AllignBottom)
                usedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
        }

        // Turn auto filter on/off
        public static void ToggleFilter()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            if (activeWorkbook == null) return;

            Excel.Worksheet activeSheet = excelApp.ActiveSheet;
            if (activeSheet == null) return;

            Excel.Range usedRange = activeSheet.UsedRange;

            if (activeSheet.AutoFilterMode)
            {
                activeSheet.AutoFilterMode = false;
            }
            else if (usedRange != null)
            {
                ApplyFilter(usedRange);
            }
        }

        // Method to normalize text/remove diacritics
        public static string NormalizeText(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // First, remove diacritics
            string normalized = input.Normalize(NormalizationForm.FormD);
            StringBuilder diacriticRemoved = new StringBuilder();

            foreach (char c in normalized)
            {
                // Keep only non-diacritic characters
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                {
                    diacriticRemoved.Append(c);
                }
            }

            // Convert to string and re-normalize
            string diacriticFreeText = diacriticRemoved.ToString().Normalize(NormalizationForm.FormC);

            // Then replace special characters
            StringBuilder result = new StringBuilder(diacriticFreeText.Length);

            foreach (char c in diacriticFreeText)
            {
                // Check if the character is in our special replacements dictionary
                if (QTUtils.specialReplacements.TryGetValue(c, out string replacement))
                {
                    result.Append(replacement);
                }
                else
                {
                    result.Append(c);
                }
            }

            return result.ToString();
        }

        // Method to replace non-ASCII characters with HTML entity
        private static string ReplaceNonAscii(string input)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();

            foreach (char c in input)
            {
                if (c <= 127)
                {
                    sb.Append(c); // Keep
                }
                else
                {
                    sb.Append($"&#{(int)c};"); // Convert
                }
            }

            return sb.ToString();
        }

        // Method to subscript Chemical formulas
        public static void SubscriptChemicalFormulas()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range rangeToProcess = QTUtils.GetRangeToProcess(excelApp);

            try
            {
                if (rangeToProcess == null)
                {
                    return;
                }

                // Create an instance of QTClipboard
                QTClipboard clipboard = QTClipboard.Instance;
                // Copy and store values
                clipboard.CopyAndStoreFormat(rangeToProcess);

                excelApp.ScreenUpdating = false;

                // Process
                foreach (Excel.Range cell in rangeToProcess.Cells)
                {
                    ConvertToSubscript(cell);
                }

            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
            finally
            {
                excelApp.ScreenUpdating = true;
                QTUtils.CleanupResources(rangeToProcess);
            }
        }

        // Method to convert numbers to subscript
        public static void ConvertToSubscript(Excel.Range cell)
        {
            if (cell == null || cell.Value2 == null)
                return;

            string formula = cell.Value2.ToString();
            char[] formulaChars = formula.ToCharArray();

            for (int i = 0; i < formulaChars.Length; i++)
            {
                if (char.IsDigit(formulaChars[i]))
                {
                    // Apply subscript formatting to the number
                    cell.Characters[i + 1, 1].Font.Subscript = true;
                }
                else
                {
                    // Ensure non-numeric characters are not subscript
                    cell.Characters[i + 1, 1].Font.Subscript = false;
                }
            }
        }

        // Reset column
        public static void ResetColumn()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;

            if (activeWorkbook != null)
            {
                Excel.Worksheet activeSheet = excelApp.ActiveSheet;
                Excel.Range selectedRange = excelApp.Selection;
                Excel.Range dataRange = null;

                try
                {
                    // Ensure the user selects only one full column
                    if (selectedRange.Columns.Count != 1 || selectedRange.Rows.Count != activeSheet.Rows.Count)
                    {
                        selectedRange = QTUtils.ColumnSelection(excelApp);

                        if (selectedRange == null)
                        {
                            return;
                        }
                    }

                    // Limit the operation to the actual rows with data
                    dataRange = selectedRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Type.Missing)
                        ?? selectedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Type.Missing)
                        ?? selectedRange;

                    // Set column to general format
                    selectedRange.NumberFormat = "General";

                    // Perform TextToColumns to reset it
                    dataRange.TextToColumns(
                        Destination: dataRange.Cells[1, 1],  // Start destination
                        DataType: Excel.XlTextParsingType.xlDelimited,
                        TextQualifier: Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
                        ConsecutiveDelimiter: false,
                        Tab: true,
                        Semicolon: false,
                        Comma: false,
                        Space: false,
                        Other: false,
                        FieldInfo: new object[,] { { 1, 1 } },
                        TrailingMinusNumbers: true
                    );

                }
                catch (Exception ex)
                {
                    QTUtils.ShowError(ex);
                }
                finally
                {
                    QTUtils.CleanupResources(selectedRange);
                    QTUtils.CleanupResources(dataRange);
                }

            }
        }
    }
}
