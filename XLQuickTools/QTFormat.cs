using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using static XLQuickTools.QTSettings;
using static XLQuickTools.QTConstants;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace XLQuickTools
{
    public class QTFormat
    {
        // Method to transform text based on an option
        private static string TransformText(string input, int option, string leading = null, string trailing = null, UserSettings settings = null)
        {

            switch (option)
            {
                case UPPERCASE:
                    return input.ToUpper();
                case LOWERCASE:
                    return input.ToLower();
                case PROPERCASE:
                    return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(input.ToLower());
                case REMOVE_LETTERS:
                    return System.Text.RegularExpressions.Regex.Replace(input, "\\p{L}", "");
                case REMOVE_NUMBERS:
                    return System.Text.RegularExpressions.Regex.Replace(input, "[0-9]", "");
                case REMOVE_SPECIAL: // Remove special characters (keep spaces and accents)
                    return System.Text.RegularExpressions.Regex.Replace(input, @"[^a-zA-Z0-9\u00C0-\u024F\s]", "");
                case NORMALIZE_TEXT: // Normalize Text/Remove diacritics
                    return NormalizeText(input);
                case TRIMCLEAN:
                    if (input != null)
                    {
                        // Trim as default
                        input = input.Trim();

                        // Remove two or more spaces
                        if (settings.Spaces)
                        {
                            input = Regex.Replace(input, @"\s{2,}", " ").Trim();
                        }

                        // Remove all spaces
                        if (settings.SpacesAll)
                        {
                            input = Regex.Replace(input, @"\s+", "").Trim();
                        }

                        // Clean
                        if (settings.NonPrintable)
                        {
                            input = Clean(input);
                        }

                        // Remove Non-ASCII
                        if (settings.NonASCII)
                        {
                            input = ReplaceNonAscii(input, false);
                        }

                        // Check for string containing only spaces
                        if (input.Length > 0 && input.All(c => c == ' '))
                        {
                            input = string.Empty;
                        }
                    }
                    return input;
                case ADD_LEADING_TRAILING:
                    return (leading ?? "") + input + (trailing ?? "");
                case REPLACE_NON_ASCII:
                    return ReplaceNonAscii(input, true);
                case REMOVE_SPACES:
                    return Regex.Replace(input, @"\s{2,}", " ").Trim();
                default:
                    throw new ArgumentException("Invalid option.");
            }
        }

        // Format select menu
        public static void FormatMenu(int option, string leading = null, string trailing = null)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range rangeToProcess = null;

            // Load settings and validate
            UserSettings settings = QTSettings.LoadUserSettingsFromXml();

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

                if (ProcessFormat(values, option, leading, trailing, settings))
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
                // Clean up
                QTUtils.CleanupResources(rangeToProcess);
            }
        }

        // Process the Format Menu option array and modify values if necessary
        private static bool ProcessFormat(object[,] values, int option, string leading = null, string trailing = null, UserSettings settings = null)
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
                            values[row, col] = TransformText(cellValue, option, leading, trailing, settings);
                            modified = true;
                        }
                    }
                }
            }
            return modified;
        }

        // Method to remove excess formatting
        public static void RemoveExcess(bool processWorkbook = false)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            Excel.Range lastRowFoundCell = null;
            Excel.Range lastColumnFoundCell = null;

            if (activeWorkbook != null)
            {
                try
                {
                    // Turn off screen updating
                    excelApp.ScreenUpdating = false;

                    // Determine which worksheets to process
                    IEnumerable<Excel.Worksheet> worksheetsToProcess = processWorkbook
                        ? activeWorkbook.Worksheets.Cast<Excel.Worksheet>()
                        : new[] { excelApp.ActiveSheet as Excel.Worksheet };

                    foreach (Excel.Worksheet ws in worksheetsToProcess)
                    {
                        // Last row with actual data (ignoring formatting)
                        lastRowFoundCell = ws.Cells.Find(
                            What: "*",
                            After: System.Reflection.Missing.Value,
                            LookIn: Excel.XlFindLookIn.xlFormulas,
                            LookAt: Excel.XlLookAt.xlPart,
                            SearchOrder: Excel.XlSearchOrder.xlByRows,
                            SearchDirection: Excel.XlSearchDirection.xlPrevious,
                            MatchCase: false,
                            MatchByte: System.Reflection.Missing.Value,
                            SearchFormat: false // Ignore formatting
                        );

                        int lastDataRow = lastRowFoundCell?.Row ?? 0;

                        // Last column with actual data (ignoring formatting)
                        lastColumnFoundCell = ws.Cells.Find(
                            What: "*",
                            After: System.Reflection.Missing.Value,
                            LookIn: Excel.XlFindLookIn.xlFormulas,
                            LookAt: Excel.XlLookAt.xlPart,
                            SearchOrder: Excel.XlSearchOrder.xlByColumns,
                            SearchDirection: Excel.XlSearchDirection.xlPrevious,
                            MatchCase: false,
                            MatchByte: System.Reflection.Missing.Value,
                            SearchFormat: false // Ignore formatting
                        );

                        int lastDataColumn = lastColumnFoundCell?.Column ?? 0;

                        // Get Excels max rows and columns allowed
                        int totalRows = ws.Rows.Count;
                        int totalColumns = ws.Columns.Count;

                        // Clear excess rows
                        if (lastDataRow < totalRows)
                        {
                            Excel.Range rowsToClear = ws.Rows[$"{lastDataRow + 1}:{totalRows}"];
                            rowsToClear.ClearFormats();
                        }

                        // Clear excess columns
                        if (lastDataColumn < totalColumns)
                        {
                            int startColumn = lastDataColumn + 1;
                            if (startColumn <= totalColumns)
                            {
                                Excel.Range columnsToClear = ws.Range[ws.Cells[1, startColumn], ws.Cells[lastDataRow + 1, totalColumns]];
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
                    // Clean up
                    QTUtils.CleanupResources(lastRowFoundCell);
                    QTUtils.CleanupResources(lastColumnFoundCell);

                    // Update message based on whether workbook or worksheet
                    string message = processWorkbook
                        ? "Excess formatting removed from workbook."
                        : "Excess formatting removed from active worksheet.";

                    System.Windows.Forms.MessageBox.Show(message,
                        "Remove Excess",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
        }

        // Method for Trim & Clean (worksheet and workbook)
        public static void TrimClean(String rangeType)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            bool processSuccess = true;

            try
            {
                excelApp.ScreenUpdating = false;

                if (rangeType.ToLower() == "worksheet")
                {
                    // Process active worksheet only
                    Excel.Worksheet activeSheet = excelApp.ActiveSheet;
                    processSuccess = ProcessSheet(activeSheet, TRIMCLEAN);
                }
                else if (rangeType.ToLower() == "workbook")
                {
                    // Process all worksheets in workbook
                    foreach (Excel.Worksheet worksheet in activeWorkbook.Worksheets)
                    {
                        if (!ProcessSheet(worksheet, TRIMCLEAN))
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
                        $"All cells in the {scope} have been trimmed and cleaned.",
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
            if (string.IsNullOrEmpty(input))
            {
                return input;
            }

            // Remove non-printable characters
            return new string(input.Where(c => !char.IsControl(c)).ToArray());

        }

        // Function to remove formatting
        public static void RemoveFormatting(Excel.Worksheet sheet)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;

            try
            {
                // Apply to the entire sheet
                Excel.Range entireSheet = sheet.Cells;

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
                if (sheet.AutoFilterMode)
                {
                    sheet.AutoFilterMode = false;
                }

                // Reset gridlines and zoom
                excelApp.ActiveWindow.DisplayGridlines = true;
                excelApp.ActiveWindow.Zoom = 100;

                // Reset row height and column width
                sheet.Rows.RowHeight = sheet.StandardHeight;
                sheet.Columns.ColumnWidth = sheet.StandardWidth;
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
        }

        // Remove formatting with screen updating turned off
        public static void RemoveFormattingNoUpdates(Excel.Worksheet activeSheet, bool applyAll = false)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            excelApp.ScreenUpdating = false;

            if (applyAll)
            {
                foreach (Excel.Worksheet sheet in excelApp.Worksheets)
                {
                    // Activate sheet
                    sheet.Activate();
                    RemoveFormatting(sheet);
                }

                // Activate original active worksheet
                activeSheet.Activate();
            }
            else
            {
                RemoveFormatting(activeSheet);
            }

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
        public static void QuickFormat(bool applyAll = false)
        {
            // Get Excel application and active workbook
            var excelApp = Globals.ThisAddIn.Application;
            var activeWorkbook = excelApp.ActiveWorkbook;
            if (activeWorkbook == null) return;

            // Active sheet
            var activeSheet = excelApp.ActiveSheet;

            // Disable screen updating
            excelApp.ScreenUpdating = false;

            try
            {
                // Load settings and validate
                UserSettings settings = QTSettings.LoadUserSettingsFromXml();
                if (settings == null) return;

                // Check if this should be applied to all worksheets
                if (applyAll)
                {
                    // Loop through all worksheets in the active workbook
                    foreach (Excel.Worksheet sheet in activeWorkbook.Worksheets)
                    {
                        // Activate each sheet
                        sheet.Activate();

                        // Get used range of the sheet
                        var usedRange = QTUtils.GetEffectiveUsedRange(sheet);
                        if (usedRange == null) continue;

                        // Remove existing formatting from the sheet
                        RemoveFormatting(sheet);

                        // Apply formatting operations to each sheet
                        ApplyRangeFormatting(usedRange, settings);
                        ApplyWindowSettings(excelApp, settings);
                        ApplyAlignmentSettings(usedRange, settings);
                        ApplyFirstRowFormatting(usedRange, settings);
                        ApplyFitFormatting(usedRange, settings);
                    }

                    // Activate active sheet
                    activeSheet.Activate();
                }
                else
                {
                    var usedRange = QTUtils.GetEffectiveUsedRange(activeSheet);
                    if (usedRange == null) return;

                    // Remove existing formatting first
                    RemoveFormatting(activeSheet);

                    // Apply formatting operations to active sheet
                    ApplyRangeFormatting(usedRange, settings);
                    ApplyWindowSettings(excelApp, settings);
                    ApplyAlignmentSettings(usedRange, settings);
                    ApplyFirstRowFormatting(usedRange, settings);
                    ApplyFitFormatting(usedRange, settings);
                }
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
        private static string ReplaceNonAscii(string input, bool html = false)
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
                    if (html)
                    {
                        sb.Append($"&#{(int)c};"); // Convert
                    }
                    else
                    {
                        // Do nothing to remove it
                    }
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

        // Method to copy column widths from source to destination
        public static void CopyColumnWidths(Excel.Range sourceRange, Excel.Range targetRange)
        {
            if (sourceRange == null) return;
            if (targetRange == null) return;

            try
            {
                // Get the number of columns in each range
                int sourceColumnCount = sourceRange.Columns.Count;
                //int targetColumnCount = targetRange.Columns.Count;

                // Iterate through each column in the source range
                for (int i = 1; i <= sourceColumnCount; i++)
                {
                    // Get the width of the current column from the source range
                    double columnWidth = (double)((Excel.Range)sourceRange.Columns[i]).ColumnWidth;
                    // Apply the width to the corresponding column in the target range
                    ((Excel.Range)targetRange.Columns[i]).ColumnWidth = columnWidth;

                }
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }

        }

        // Method to copy active worksheet formatting to all worksheets
        public static void CopyFormatToAllSheets()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet;

            // Copy the used range of the active sheet
            Excel.Range sourceRange = activeSheet.UsedRange;

            try
            {
                if (workbook == null || activeSheet == null)
                {
                    return;
                }

                excelApp.ScreenUpdating = false;

                // Loop through all sheets in the workbook
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    // Skip the active sheet
                    if (sheet.Name != activeSheet.Name)
                    {
                        // Get the target range of the same size as source range
                        Excel.Range targetRange = sheet.Range[sourceRange.Address];

                        // Copy and apply the formatting
                        sourceRange.Copy();
                        targetRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);

                        // Copy column widths
                        CopyColumnWidths(sourceRange, targetRange);

                        // Move the cursor to cell A1
                        sheet.Activate();
                        sheet.Range["A1"].Select();

                    }
                }

                // Go back to main sheet
                activeSheet.Activate();
                activeSheet.Range["A1"].Select();

                // Clear the clipboard
                excelApp.CutCopyMode = 0;
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
            finally
            {
                excelApp.ScreenUpdating = true;
                QTUtils.CleanupResources(sourceRange);
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
                Excel.Range columnRange = null;

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

                    columnRange = selectedRange.EntireColumn;

                    // Set to General format
                    columnRange.NumberFormat = "General";

                    // Apply TextToColumns to the whole column
                    columnRange.TextToColumns(
                        Destination: columnRange.Cells[1, 1],
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
                    QTUtils.CleanupResources(columnRange);
                    QTUtils.CleanupResources(selectedRange);
                }
            }
        }
    }
}

