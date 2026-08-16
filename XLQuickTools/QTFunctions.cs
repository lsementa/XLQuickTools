using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using static XLQuickTools.QTConstants;
using static XLQuickTools.QTSettings;
using static XLQuickTools.QTUtils;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    internal class QTFunctions
    {
        private const int HeaderRow = 1;
        private const int FirstDataRow = 2;
        private const bool TreatNumericTextAsNumber = true;

        // Selection and Selection+ function
        public static void SelectionPlus(string leading = "", string trailing = "", string delimiter = ",", int newLine = 0)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range rangeToProcess = QTUtils.GetRangeToProcess(excelApp);
            if (rangeToProcess == null) return;

            try
            {
                StringBuilder sb = new StringBuilder();
                bool firstValue = true;

                // Set the range to an array
                var values = QTUtils.GetRangeValues(rangeToProcess);
                if (values == null) return;

                int baseIndex = values.GetLowerBound(0);
                int rowsCount = values.GetUpperBound(0);
                int colsCount = values.GetUpperBound(1);

                for (int row = baseIndex; row <= rowsCount; row++)
                {
                    for (int col = baseIndex; col <= colsCount; col++)
                    {
                        // Convert cell value to string and trim whitespace
                        string cellValue = values[row, col]?.ToString()?.Trim();

                        // Only append if cellValue is not empty or null
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            // Append comma only if this is not the first value and no newline flag
                            if (!firstValue && newLine == 0)
                            {
                                sb.Append(delimiter + " ");
                            }

                            // Append leading and trailing character
                            sb.Append(leading + cellValue + trailing);

                            // Add a newline if the `newLine` flag is set
                            if (newLine == 1)
                            {
                                sb.Append(delimiter + "\n");
                            }
                            else if (newLine == 2)
                            {
                                sb.Append("\n");
                            }

                            firstValue = false;
                        }
                    }
                }

                // If newline is true, remove the last newline character
                if ((newLine == 1 || newLine == 2) && sb.Length > 0)
                {
                    sb.Length--; // Removes the last appended newline

                    if (newLine == 1)
                    {
                        sb.Length--; // Additional removal for the extra comma
                    }
                }

                // Copy to clipboard if any non-empty values were found
                if (sb.Length > 0)
                {
                    System.Windows.Forms.Clipboard.SetText(sb.ToString());
                }
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
            finally
            {
                excelApp.StatusBar = "";
                QTUtils.CleanupResources(rangeToProcess);
            }
        }

        // Method to fill blanks based on the Excel range
        public static void FillBlanks(Excel.Range fillRange)
        {
            try
            {
                // Validate the input range
                if (fillRange == null) return;

                // Try to get blank cells in the range
                Excel.Range blanks;
                try
                {
                    blanks = fillRange.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    return;
                }

                // Iterate through each blank cell and fill it with the value from above
                foreach (Excel.Range blank in blanks)
                {
                    Excel.Range aboveCell = blank.Offset[-1, 0];  // Get the cell above
                    blank.Value2 = aboveCell.Value2;  // Copy value from above
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Invalid selection. Can't fill blanks from above with the current range.", "Invalid Selection",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Method for Fill Down
        public static void FillDown()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range rangeToProcess = QTUtils.GetRangeToProcess(excelApp);
            if (rangeToProcess == null) return;

            // Create an instance of QTClipboard
            QTClipboard clipboard = QTClipboard.Instance;
            // Copy and store values
            clipboard.CopyAndStoreFormat(rangeToProcess);

            FillBlanks(rangeToProcess);
        }

        // Method to copy unique values to new worksheets
        public static void CopyToSheets()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;

            if (activeWorkbook != null)
            {
                Excel.Worksheet activeSheet = excelApp.ActiveSheet;
                Excel.Range usedRange = activeSheet.UsedRange;
                Excel.Range selectedRange = excelApp.Selection;

                try
                {

                    // Ensure the user selects only one full column
                    if (selectedRange.Columns.Count != 1 || selectedRange.Rows.Count != activeSheet.Rows.Count)
                    {
                        Excel.Range selectedColumn = QTUtils.ColumnSelection(excelApp);

                        if (selectedColumn != null)
                        {
                            selectedRange = selectedColumn;
                        }
                        else
                        {
                            return;
                        }
                    }

                    // Ensure the column contains a header
                    if (selectedRange.Cells[1, 1].Value == null || string.IsNullOrWhiteSpace(selectedRange.Cells[1, 1].Value.ToString()))
                    {
                        MessageBox.Show("Column must contain a header.", "Invalid Selection",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Ensure Column A contains a header for it to be a complete worksheet
                    if (activeSheet.Cells[1, 1].Value == null || string.IsNullOrWhiteSpace(activeSheet.Cells[1, 1].Value.ToString()))
                    {
                        MessageBox.Show("No heading in column A.", "Incomplete Worksheet",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Use AutoFilter to filter unique criteria
                    Excel.Range columnRange = selectedRange.Columns[1];
                    int lastRow = activeSheet.Cells[activeSheet.Rows.Count, columnRange.Column].End(Excel.XlDirection.xlUp).Row;

                    // Nothing below the header row
                    if (lastRow < 2)
                    {
                        MessageBox.Show("No data found below the header.", "Nothing to Process",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    Excel.Range dataRange = activeSheet.Range[activeSheet.Cells[1, 1], activeSheet.Cells[lastRow, activeSheet.Columns.Count]];

                    // Read the column in one pass rather than cell by cell
                    object[,] columnValues = activeSheet.Range[
                        activeSheet.Cells[1, columnRange.Column],
                        activeSheet.Cells[lastRow, columnRange.Column]].Value as object[,];

                    // Get unique values from the column
                    var uniqueValues = new HashSet<object>();
                    if (columnValues != null)
                    {
                        for (int row = 2; row <= lastRow; row++) // Skip the header row
                        {
                            object value = columnValues[row, 1];
                            if (value != null && !uniqueValues.Contains(value))
                            {
                                uniqueValues.Add(value);
                            }
                        }
                    }

                    // Show the count up front. The user decides whether to proceed
                    using (UniqueSheetsForm form1 = new UniqueSheetsForm(selectedRange, uniqueValues.Count))
                    {
                        if (form1.ShowDialog() != DialogResult.OK) return;
                    }

                    // Turn off screen updating
                    excelApp.ScreenUpdating = false;

                    Excel.Worksheet lastSheet = activeSheet; // Start with the active sheet

                    // Iterate through each unique value and copy filtered data to a new sheet
                    foreach (object uniqueValue in uniqueValues)
                    {
                        string sheetName = uniqueValue.ToString();

                        // Check if a sheet with the same name already exists
                        bool sheetExists = false;
                        foreach (Excel.Worksheet sheet in activeWorkbook.Sheets)
                        {
                            if (sheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                sheetExists = true;
                                break;
                            }
                        }

                        if (sheetExists)
                        {
                            // Display error message and exit the loop
                            MessageBox.Show($"A sheet with the name '{sheetName}' already exists. Aborting operation.", "Worksheet", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            break;
                        }

                        // Filter the data based on the unique value
                        dataRange.AutoFilter(Field: columnRange.Column, Criteria1: uniqueValue);

                        // Create a new worksheet, placing it after the last created sheet
                        Excel.Worksheet newSheet = activeWorkbook.Sheets.Add(After: lastSheet);
                        newSheet.Name = sheetName;

                        // Copy visible data (filtered) to the new sheet
                        dataRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy(newSheet.Cells[1, 1]);

                        // Copy column widths
                        Excel.Range newUsedRange = newSheet.UsedRange;
                        QTFormat.CopyColumnWidths(usedRange, newUsedRange);

                        // Update lastSheet to the newly created sheet
                        lastSheet = newSheet;
                    }

                    // Remove AutoFilter
                    activeSheet.AutoFilterMode = false;
                }
                finally
                {
                    // Turn screen updating back on
                    excelApp.ScreenUpdating = true;
                    QTUtils.CleanupResources(usedRange);
                }
            }
        }

        // Find duplicates - acts as a toggle. First click adds the Count column and
        // stores the location, second click returns to it and removes the column.
        public static void FindDuplicates()
        {
            QTClipboard clipboard = QTClipboard.Instance;

            // Toggle OFF - a Count column created this session is still live somewhere.
            // Handled before touching the selection, since the user may have moved away.
            if (clipboard.HasDuplicatesState)
            {
                RemoveDuplicatesCount();
                return;
            }

            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;

            if (activeWorkbook != null)
            {
                Excel.Worksheet activeSheet = excelApp.ActiveSheet;
                Excel.Range selectedRange = excelApp.Selection;

                try
                {
                    // Ensure the user selects only one full column
                    if (selectedRange.Columns.Count != 1 || selectedRange.Rows.Count != activeSheet.Rows.Count)
                    {
                        Excel.Range selectedColumn = QTUtils.ColumnSelection(excelApp);

                        if (selectedColumn != null)
                        {
                            selectedRange = selectedColumn;
                        }
                        else
                        {
                            return;
                        }
                    }

                    // Ensure the column contains a header
                    if (selectedRange.Cells[1, 1].Value == null || string.IsNullOrWhiteSpace(selectedRange.Cells[1, 1].Value.ToString()))
                    {
                        MessageBox.Show("Column must contain a header.", "Invalid Selection",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Turn off screen updating
                    excelApp.ScreenUpdating = false;

                    // Get column to the right
                    int nextColumn = selectedRange.Column + 1;

                    // Reference the header of the adjacent column
                    Excel.Range adjacentHeaderCell = activeSheet.Cells[1, nextColumn];

                    // Fallback for a Count column left over from a previous session
                    // (add-in reload, Excel restart) where the stored state is gone
                    if (adjacentHeaderCell.Value != null && adjacentHeaderCell.Value.ToString() == "Count")
                    {
                        // Remove Autofilter if applied
                        if (activeSheet.AutoFilterMode) activeSheet.AutoFilterMode = false;

                        // Delete the "Count" column
                        adjacentHeaderCell.EntireColumn.Delete();

                        // No state to clear, but this resets the button label
                        clipboard.ClearDuplicatesState();
                        return;
                    }

                    // Get the selected column range
                    Excel.Range columnRange = selectedRange.Columns[1];

                    // Read the values from the selected range into an array
                    // (1-based, and safe when the column holds a single cell)
                    object[,] columnValues = QTUtils.GetValueArray(columnRange);
                    if (columnValues == null) return;

                    // Create a dictionary to track the count of each unique value
                    Dictionary<string, int> valueCount = new Dictionary<string, int>();
                    bool hasDuplicates = false;

                    // Iterate through the array to find duplicates and count occurrences
                    for (int i = 2; i <= columnValues.GetLength(0); i++) // Start from row 2 to skip header
                    {
                        object cellValueObj = columnValues[i, 1]; // Read the value from the array

                        if (cellValueObj != null && !string.IsNullOrWhiteSpace(cellValueObj.ToString())) // Ignore blanks
                        {
                            string cellValue = cellValueObj.ToString();

                            if (valueCount.ContainsKey(cellValue))
                            {
                                valueCount[cellValue]++;
                                hasDuplicates = true; // Mark that there are duplicates
                            }
                            else
                            {
                                valueCount[cellValue] = 1;
                            }
                        }
                    }

                    // If duplicates are found, insert a new column with the counts
                    if (hasDuplicates)
                    {
                        // Insert a new column to the right of the selected column
                        Excel.Range newColumn = columnRange.Offset[0, 1].EntireColumn;
                        newColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                        // Prepare the count array for the new column
                        object[,] countValues = new object[columnValues.GetLength(0), 1];

                        // Populate the array with the count of each value
                        for (int i = 2; i <= columnValues.GetLength(0); i++) // Start from row 2 to skip header
                        {
                            object cellValueObj = columnValues[i, 1];

                            if (cellValueObj != null && !string.IsNullOrWhiteSpace(cellValueObj.ToString()))
                            {
                                string cellValue = cellValueObj.ToString();
                                countValues[i - 1, 0] = valueCount[cellValue]; // Fill the count
                            }
                        }

                        // Write the count array back to the worksheet in one operation
                        Excel.Range countRange = activeSheet.Range[activeSheet.Cells[1, nextColumn], activeSheet.Cells[columnValues.GetLength(0), nextColumn]];
                        countRange.Value2 = countValues;

                        // Set the header of the new column to "Count"
                        activeSheet.Cells[1, nextColumn].Value = "Count";

                        // Apply AutoFilter to the count column
                        if (activeSheet.AutoFilterMode)
                        {
                            activeSheet.AutoFilterMode = false; // Clear any existing AutoFilter
                        }

                        // Define the range that includes all headers and data
                        Excel.Range fullDataRange = activeSheet.UsedRange; // Covers all non-empty cells in the worksheet

                        // Apply AutoFilter to the "Count" column (>1)
                        int relativeCountColumn = nextColumn; // Use the actual column index for filtering

                        // Apply AutoFilter to the full range
                        try
                        {
                            fullDataRange.AutoFilter(relativeCountColumn, ">1", Excel.XlAutoFilterOperator.xlAnd);
                        }
                        catch (Exception ex)
                        {
                            QTFormat.ApplyFilter(fullDataRange);
                            MessageBox.Show($"Error applying AutoFilter: {ex.Message}\nCount column has been left unfilterd.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }

                        // Store the location last, so a failure above leaves the button in the off state
                        clipboard.StoreDuplicatesState(columnRange, nextColumn);
                    }
                    else
                    {
                        MessageBox.Show("No duplicates found in the selected column.", "No Duplicates",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                }
            }
        }

        // Return to the stored selection, drop the Count column and reset the state
        private static void RemoveDuplicatesCount()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            QTClipboard clipboard = QTClipboard.Instance;

            Excel.Worksheet worksheet;
            Excel.Range sourceRange;
            int countColumn;

            // Workbook or sheet is gone - nothing to undo, just reset the button
            if (!clipboard.TryGetDuplicatesTarget(out worksheet, out sourceRange, out countColumn))
            {
                clipboard.ClearDuplicatesState();
                return;
            }

            try
            {
                excelApp.ScreenUpdating = false;

                // Bring the stored selection back into view
                ((Excel.Workbook)worksheet.Parent).Activate();
                worksheet.Activate();

                // Remove the filter that was applied to the Count column
                if (worksheet.AutoFilterMode) worksheet.AutoFilterMode = false;

                // Only delete if it is still the Count column
                Excel.Range header = worksheet.Cells[1, countColumn];
                if (header.Value2 != null &&
                    string.Equals(header.Value2.ToString(), "Count", StringComparison.OrdinalIgnoreCase))
                {
                    header.EntireColumn.Delete();
                }

                sourceRange.Select();
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
            finally
            {
                excelApp.ScreenUpdating = true;
                clipboard.ClearDuplicatesState();
                QTUtils.CleanupResources(sourceRange);
            }
        }

        // Remove hyperlinks
        public static void RemoveHyperlinks()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;

            if (activeWorkbook != null)
            {
                Excel.Worksheet activeSheet = excelApp.ActiveSheet;
                Excel.Range selectedRange = excelApp.Selection;

                // Ensure the user selects only one full column
                if (selectedRange.Columns.Count != 1 || selectedRange.Rows.Count != activeSheet.Rows.Count)
                {
                    Excel.Range selectedColumn = QTUtils.ColumnSelection(excelApp);

                    if (selectedColumn != null)
                    {
                        selectedRange = selectedColumn;
                    }
                    else
                    {
                        return;
                    }
                }
                try
                {
                    excelApp.ScreenUpdating = false;
                    // Set to Text
                    selectedRange.NumberFormat = "@";
                    selectedRange.Value2 = selectedRange.Value2;
                    // Remove hyperlink formatting: Reset font color and underline
                    selectedRange.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone;
                    selectedRange.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    // Bulk remove cell-based hyperlinks
                    selectedRange.Hyperlinks.Delete();
                }
                finally
                {
                    // Turn screen updating back on
                    excelApp.ScreenUpdating = true;
                }
            }
        }

        // Add hyperlinks
        public static void AddHyperlinks(bool custom)
        {
            UserSettings settings = LoadUserSettingsFromXml();
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            string selectedURL = string.Empty;

            // Find the "use this" hyperlink
            var matchingEntry = settings.HyperlinkEntries
                .FirstOrDefault(entry => entry.Use == true);

            // If it exists set all the fields
            if (matchingEntry != null)
            {
                selectedURL = matchingEntry.URL;
            }

            if (activeWorkbook != null)
            {
                Excel.Worksheet activeSheet = excelApp.ActiveSheet;
                Excel.Range selectedRange = excelApp.Selection;

                // Ensure the user selects only one full column
                if (selectedRange.Columns.Count != 1 || selectedRange.Rows.Count != activeSheet.Rows.Count)
                {
                    Excel.Range selectedColumn = QTUtils.ColumnSelection(excelApp);

                    if (selectedColumn != null)
                    {
                        selectedRange = selectedColumn;
                    }
                    else
                    {
                        return;
                    }
                }
                try
                {
                    excelApp.ScreenUpdating = false;

                    // Process the column
                    Excel.Range lastCell = selectedRange.Cells[selectedRange.Rows.Count, 1].End(Excel.XlDirection.xlUp);
                    int lastRow = lastCell.Row;

                    // Read the column values into an array
                    // (1-based, and safe when the selection is a single cell)
                    object[,] values = QTUtils.GetValueArray(selectedRange);
                    if (values == null) return;

                    int totalRows = lastRow;

                    // Process the data in chunks
                    for (int startRow = 2; startRow <= totalRows; startRow += CHUNK_SIZE)
                    {
                        int endRow = Math.Min(startRow + CHUNK_SIZE - 1, totalRows);
                        int rowsToProcess = endRow - startRow + 1;

                        object[,] processArray = new object[rowsToProcess, 1];

                        // Copy data into a new array, excluding the header
                        for (int i = startRow; i <= endRow; i++)
                        {
                            processArray[i - startRow, 0] = values[i, 1]; // Skip header
                        }

                        // Iterate through the processed chunk
                        for (int i = 0; i < rowsToProcess; i++)
                        {
                            string cellValue = processArray[i, 0]?.ToString() ?? string.Empty;
                            string hyperlinkURL = "";
                            string hyperlinkFormula = "";

                            // Create a hyperlink
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                if (custom)
                                {
                                    hyperlinkURL = selectedURL.Replace("{ID}", cellValue).Replace("{id}", cellValue);
                                    hyperlinkFormula = $"=HYPERLINK(\"{hyperlinkURL}\", \"{cellValue}\")";
                                }
                                else
                                {
                                    hyperlinkURL = cellValue;

                                    // Ensure the URL has a protocol for the HYPERLINK function to work reliably
                                    if (!hyperlinkURL.StartsWith("http://") && !hyperlinkURL.StartsWith("https://"))
                                    {
                                        hyperlinkURL = "https://" + hyperlinkURL;
                                    }

                                    hyperlinkFormula = $"=HYPERLINK(\"{hyperlinkURL}\", \"{cellValue}\")";
                                }

                                processArray[i, 0] = hyperlinkFormula;
                            }
                        }

                        // Set column to General format
                        selectedRange.NumberFormat = "General";

                        // Write the processed array back to Excel
                        Excel.Range rangeToUpdate = selectedRange.Cells[startRow, 1].Resize[rowsToProcess, 1];
                        rangeToUpdate.Value2 = processArray;
                    }
                }
                finally
                {
                    // Turn screen updating back on
                    excelApp.ScreenUpdating = true;
                }
            }
        }

        // Method to remove empty rows or columns in the active sheet
        public static void DeleteEmptyRowsOrColumns(DeleteOption option)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            if (excelApp == null)
                throw new ArgumentNullException(nameof(excelApp));

            Excel.Worksheet activeSheet = excelApp.ActiveSheet;

            // Get the used range in the active worksheet
            Excel.Range usedRange = activeSheet.UsedRange;

            // Check if the worksheet is effectively empty
            if (usedRange == null || (usedRange.Cells.Count == 1 && usedRange.Value2 == null))
            {
                return;
            }

            // Get the last row and column
            int lastRow = usedRange.Row + usedRange.Rows.Count - 1;
            int lastColumn = usedRange.Column + usedRange.Columns.Count - 1;

            // Turn off screen updating
            excelApp.ScreenUpdating = false;

            int deleteCount = 0;
            string deleteType = null;

            try
            {
                if (option == DeleteOption.Rows)
                {
                    deleteType = "row";
                    deleteCount = QTUtils.DeleteRows(activeSheet, excelApp, lastRow, usedRange.Row);
                }
                else if (option == DeleteOption.Columns)
                {
                    deleteType = "Column";
                    deleteCount = QTUtils.DeleteColumns(activeSheet, excelApp, lastColumn, usedRange.Column, lastRow);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error cleaning up {option}: {ex.Message}", ex);
            }
            finally
            {
                // Restore screen updating
                excelApp.ScreenUpdating = true;
                ShowMessage(deleteType, deleteCount);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Helper method to display a success message for deleting rows or columns
        private static void ShowMessage(string type, int count)
        {
            string message = count > 0
                ? $"{count} {type}(s) deleted from the active worksheet."
                : $"No empty {type}s were found.";

            System.Windows.Forms.MessageBox.Show(message,
                $"Delete Empty {char.ToUpper(type[0]) + type.Substring(1)}s",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        // Split a cell value on the delimiter. Returns the original value untouched
        // when no delimiter is in play, so non-delimited behavior is unchanged.
        private static List<string> SplitValue(string value, string delimiter, bool useDelimiter)
        {
            List<string> result = new List<string>();

            if (!useDelimiter || string.IsNullOrEmpty(value))
            {
                result.Add(value ?? string.Empty);
                return result;
            }

            // Delimiter is set but this cell does not contain it
            if (value.IndexOf(delimiter, StringComparison.Ordinal) < 0)
            {
                result.Add(value.Trim());
                return result;
            }

            foreach (string part in value.Split(new string[] { delimiter }, StringSplitOptions.None))
            {
                string trimmed = part.Trim();
                if (!string.IsNullOrEmpty(trimmed))
                {
                    result.Add(trimmed);
                }
            }

            // Value was nothing but delimiters
            if (result.Count == 0) result.Add(string.Empty);

            return result;
        }

        // Cartesian expansion across the checked columns. With no delimiter every
        // list holds a single entry, so this yields exactly one combination per row.
        private static IEnumerable<string[]> ExpandCombinations(List<List<string>> tokens)
        {
            int count = tokens.Count;
            if (count == 0) yield break;

            int[] idx = new int[count];

            while (true)
            {
                string[] combo = new string[count];
                for (int i = 0; i < count; i++)
                {
                    combo[i] = tokens[i][idx[i]];
                }

                yield return combo;

                // Advance the odometer
                int pos = count - 1;
                while (pos >= 0)
                {
                    idx[pos]++;
                    if (idx[pos] < tokens[pos].Count) break;
                    idx[pos] = 0;
                    pos--;
                }

                if (pos < 0) yield break;
            }
        }

        // Get the Unique count
        public static int GetUniqueCount(Excel.Range rangeToProcess, string delimiter = "")
        {
            try
            {
                if (rangeToProcess == null) return 0;

                var values = QTUtils.GetRangeValues(rangeToProcess);
                if (values == null) return 0;

                bool useDelimiter = !string.IsNullOrEmpty(delimiter);

                // Flatten the values into a single list, excluding null or empty cells,
                // splitting each cell on the delimiter when one is set
                var allValues = values.Cast<object>()
                                      .Where(v => v != null && !string.IsNullOrWhiteSpace(v.ToString()))
                                      .SelectMany(v => SplitValue(v.ToString(), delimiter, useDelimiter))
                                      .Where(s => !string.IsNullOrWhiteSpace(s))
                                      .ToList();

                // Unique count
                int uniqueCount = allValues.Distinct().Count();

                return uniqueCount;
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
                return 0; // Return a default value in case of an error
            }
        }

        // Get the unique row count with optional clipboard copy.
        // skipFirstRow treats row 1 as a header: it is excluded from the uniqueness
        // comparison and written to the clipboard verbatim as the first line.
        public static int GetUniqueRows(Excel.Range rangeToProcess, CheckedListBox clbColumns,
                                        bool copyToClipboard = false, string delimiter = "",
                                        bool skipFirstRow = false)
        {
            try
            {
                if (rangeToProcess == null || clbColumns.CheckedItems.Count == 0)
                    return 0; // No selection, return 0

                // Get the indices of checked columns
                List<int> checkedColumnIndices = new List<int>();
                foreach (var item in clbColumns.CheckedItems)
                {
                    int index = clbColumns.Items.IndexOf(item);
                    if (index >= 0)
                        checkedColumnIndices.Add(index);
                }

                // Dictionary to track unique combinations with their original row data
                Dictionary<string, List<object>> uniqueRowData = new Dictionary<string, List<object>>();

                // Convert range to array for faster processing
                object[,] data = QTUtils.GetValueArray(rangeToProcess);
                if (data == null) return 0;

                int rowCount = data.GetLength(0);
                int colCount = data.GetLength(1);

                bool useDelimiter = !string.IsNullOrEmpty(delimiter);

                // Start below the header row when one is present
                int startRow = skipFirstRow ? 2 : 1;

                // Process each row
                for (int row = startRow; row <= rowCount; row++)
                {
                    // Build the token lists for the checked columns on this row
                    List<int> activeColumns = new List<int>();
                    List<List<string>> rowTokens = new List<List<string>>();
                    bool rowHasValue = false;

                    foreach (int colIndex in checkedColumnIndices)
                    {
                        // Check if the column index is within the range
                        if (colIndex >= colCount) continue;

                        // Get column value
                        object value = data[row, colIndex + 1]; // +1 because Excel arrays are 1-based
                        string strValue = value?.ToString() ?? string.Empty;

                        // Check if this cell has a value
                        if (!string.IsNullOrWhiteSpace(strValue))
                        {
                            rowHasValue = true;
                        }

                        activeColumns.Add(colIndex);
                        rowTokens.Add(SplitValue(strValue, delimiter, useDelimiter));
                    }

                    // Only add non-blank rows to the unique set
                    if (!rowHasValue) continue;

                    // One pass per delimited value (a single pass when no delimiter is set)
                    foreach (string[] combo in ExpandCombinations(rowTokens))
                    {
                        // Build a key string from the checked columns for this combination
                        StringBuilder keyBuilder = new StringBuilder();
                        foreach (string token in combo)
                        {
                            keyBuilder.Append(token);
                            keyBuilder.Append("|"); // Use a separator unlikely to appear in cell values
                        }

                        string key = keyBuilder.ToString();
                        if (uniqueRowData.ContainsKey(key)) continue;

                        // If we're copying to clipboard, store the entire row
                        if (copyToClipboard)
                        {
                            // Extract all cell values for this row
                            List<object> rowValues = new List<object>();
                            for (int col = 1; col <= colCount; col++)
                            {
                                rowValues.Add(data[row, col]);
                            }

                            // Replace the checked columns with the individual split value
                            for (int i = 0; i < activeColumns.Count; i++)
                            {
                                rowValues[activeColumns[i]] = combo[i];
                            }

                            uniqueRowData.Add(key, rowValues);
                        }
                        else
                        {
                            // If not copying, just track the unique keys
                            uniqueRowData.Add(key, null); // Just use null as we don't need the data
                        }
                    }
                }

                // Copy to clipboard if requested
                if (copyToClipboard && uniqueRowData.Count > 0)
                {
                    // Create a string builder for the clipboard text
                    StringBuilder clipboardContent = new StringBuilder();

                    // Write the header row first, exactly as it appears on the sheet
                    if (skipFirstRow && rowCount >= 1)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            clipboardContent.Append(data[1, col]?.ToString() ?? string.Empty);

                            if (col < colCount)
                            {
                                clipboardContent.Append("\t");
                            }
                        }

                        clipboardContent.AppendLine();
                    }

                    // Add each unique row to the clipboard content
                    foreach (var kvp in uniqueRowData)
                    {
                        if (kvp.Value != null) // Should always be non-null when copyToClipboard is true
                        {
                            for (int i = 0; i < kvp.Value.Count; i++)
                            {
                                clipboardContent.Append(kvp.Value[i]?.ToString() ?? string.Empty);

                                // Add tab separator between cells, but not after the last cell
                                if (i < kvp.Value.Count - 1)
                                {
                                    clipboardContent.Append("\t");
                                }
                            }

                            // Add newline after each row
                            clipboardContent.AppendLine();
                        }
                    }

                    // Copy to clipboard
                    System.Windows.Forms.Clipboard.SetText(clipboardContent.ToString());
                }

                return uniqueRowData.Count;
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
                return 0;
            }
        }

        // Unique select options for count or copying to clipboard
        public static void UniqueSelect()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;
            Excel.Range rangeToProcess = null;

            if (activeSheet == null)
            {
                MessageBox.Show("Please select a worksheet.", "Unique Select",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                rangeToProcess = QTUtils.GetRangeToProcess(excelApp);
                if (rangeToProcess == null) return;

                // Single cell: widen to CurrentRegion, then UsedRange
                if (rangeToProcess.Cells.Count == 1)
                {
                    rangeToProcess = ReplaceRange(rangeToProcess, rangeToProcess.CurrentRegion);

                    if (rangeToProcess.Cells.Count == 1)
                    {
                        rangeToProcess = ReplaceRange(rangeToProcess, activeSheet.UsedRange);
                    }
                }

                // Trim whole-column / whole-row selections down to actual data
                Excel.Range used = activeSheet.UsedRange;
                if (rangeToProcess.Cells.Count > used.Cells.Count)
                {
                    Excel.Range trimmed = excelApp.Intersect(rangeToProcess, used);
                    if (trimmed != null)
                    {
                        rangeToProcess = ReplaceRange(rangeToProcess, trimmed);
                    }
                }

                // Test for real emptiness
                if (Convert.ToDouble(excelApp.WorksheetFunction.CountA(rangeToProcess)) == 0)
                {
                    MessageBox.Show("No data found to process.", "Unique Select",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (UniqueDataForm form1 = new UniqueDataForm(rangeToProcess))
                {
                    form1.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
            finally
            {
                QTUtils.CleanupResources(rangeToProcess);
            }
        }

        // Swap in a new range and release the old one
        private static Excel.Range ReplaceRange(Excel.Range oldRange, Excel.Range newRange)
        {
            if (!ReferenceEquals(oldRange, newRange))
            {
                QTUtils.CleanupResources(oldRange);
            }
            return newRange;
        }

        // Sheet names to clipboard
        public static void CopyWorksheetNamesToClipboard()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            if (activeWorkbook == null) return;

            StringBuilder worksheetNames = new StringBuilder();

            // Iterate through each sheet and create the list
            foreach (Excel.Worksheet sheet in activeWorkbook.Worksheets)
            {
                worksheetNames.AppendLine(sheet.Name);
            }

            // Copy to clipboard
            if (worksheetNames.Length > 0)
            {
                Clipboard.SetText(worksheetNames.ToString());
            }

        }

        // Copy highlighted cells to clipboard
        public static void CopyHighlightedCellsToClipboard()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet;
            Excel.Range rangeToProcess = null;

            try
            {
                rangeToProcess = QTUtils.GetRangeToProcess(excelApp);
                if (rangeToProcess == null) return;

                StringBuilder clipboardText = new StringBuilder();

                int rowCount = rangeToProcess.Rows.Count;
                int colCount = rangeToProcess.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        Excel.Range cell = rangeToProcess.Cells[i, j];

                        // Check if the cell's interior color is not the default 'no fill' color.
                        if (cell.Interior.ColorIndex != (int)Excel.XlColorIndex.xlColorIndexNone)
                        {
                            // Text instead of cell.value2
                            string displayText = cell.Text?.ToString();
                            if (!string.IsNullOrWhiteSpace(displayText))
                            {
                                clipboardText.AppendLine(displayText);
                            }
                        }
                    }
                }

                if (clipboardText.Length > 0)
                {
                    Clipboard.SetText(clipboardText.ToString());
                }
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
            finally
            {
                QTUtils.CleanupResources(rangeToProcess);
            }
        }

        // Column stats used in Column Information
        public class ColumnStats
        {
            public string ColumnLetter { get; set; } = string.Empty;
            public string ColumnName { get; set; } = string.Empty;
            public int UniqueValues { get; set; }
            public int DuplicateValues { get; set; }
            public int NonBlankCells { get; set; }
            public int BlankCells { get; set; }
            public int RowCount { get; set; }
        }

        // Trim a full column down to the rows actually in use
        private static Excel.Range GetColumnDataRange(Excel.Worksheet sheet, Excel.Range column)
        {
            if (sheet == null || column == null) return null;

            Excel.Range usedRange = sheet.UsedRange;
            Excel.Range intersect = sheet.Application.Intersect(column.EntireColumn, usedRange);

            return intersect;
        }

        // Build the counts for a single column
        public static ColumnStats GetColumnStats(Excel.Range columnRange, bool hasHeaders)
        {
            ColumnStats stats = new ColumnStats();
            if (columnRange == null) return stats;

            stats.ColumnLetter = QTUtils.GetColumnLetter(columnRange.Column);
            stats.ColumnName = "Column [ " + stats.ColumnLetter + " ]";

            // Handles both a multi cell range and a single cell
            object[,] data = QTUtils.GetValueArray(columnRange);

            // Tracks each distinct value and how many times it occurs
            Dictionary<string, int> valueCounts = new Dictionary<string, int>();
            int totalRows = 0;
            int nonBlank = 0;

            if (data != null)
            {
                int firstRow = data.GetLowerBound(0);
                int lastRow = data.GetUpperBound(0);
                int col = data.GetLowerBound(1);

                // Use row 1 as the column name
                if (hasHeaders)
                {
                    object header = data[firstRow, col];
                    string headerText = header == null ? string.Empty : header.ToString().Trim();
                    if (headerText.Length > 0) stats.ColumnName = headerText;
                }

                int startRow = hasHeaders ? firstRow + 1 : firstRow;

                for (int r = startRow; r <= lastRow; r++)
                {
                    totalRows++;

                    object value = data[r, col];
                    string text = value == null ? string.Empty : value.ToString().Trim();
                    if (text.Length == 0) continue;

                    nonBlank++;

                    if (valueCounts.ContainsKey(text))
                        valueCounts[text]++;
                    else
                        valueCounts[text] = 1;
                }
            }

            stats.RowCount = totalRows;
            stats.NonBlankCells = nonBlank;
            stats.BlankCells = totalRows - nonBlank;
            stats.UniqueValues = valueCounts.Count;

            // Distinct values that occur more than once
            stats.DuplicateValues = valueCounts.Count(kv => kv.Value > 1);

            return stats;
        }

        // Column information
        public static void CountValuesInColumn()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            if (activeWorkbook == null) return;

            Excel.Worksheet activeSheet = excelApp.ActiveSheet;
            Excel.Range selectedRange = excelApp.Selection;
            Excel.Range columnRange = null;

            try
            {
                // Ensure the user selects only one full column
                if (selectedRange.Columns.Count != 1 || selectedRange.Rows.Count != activeSheet.Rows.Count)
                {
                    Excel.Range selectedColumn = QTUtils.ColumnSelection(excelApp);

                    if (selectedColumn != null)
                    {
                        selectedRange = selectedColumn;
                    }
                    else
                    {
                        return;
                    }
                }

                // Limit the column to the used rows
                columnRange = GetColumnDataRange(activeSheet, selectedRange);
                if (columnRange == null)
                {
                    MessageBox.Show("The selected column is empty.", "Column Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Get all the values
                ColumnStats stats = GetColumnStats(columnRange, false);

                // Show all the values
                string message = $"Column [ {stats.ColumnLetter} ]\n" +
                            "──────────────────────────────\n" +
                            $" Unique Values:\t{stats.UniqueValues:N0}\n" +
                            $" Duplicate Values:\t{stats.DuplicateValues:N0}\n" +
                            $" Non-Blank Cells:\t{stats.NonBlankCells:N0}\n" +
                            $" Blank Cells:\t{stats.BlankCells:N0}\n" +
                            $" Number of Rows:\t{stats.RowCount:N0}\n" +
                            "──────────────────────────────\n";
                MessageBox.Show(message, "Column Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        // Compare two columns and insert true/false column to the right of each
        public static void CompareColumns()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            if (excelApp == null || excelApp.ActiveWorkbook == null)
                return;

            bool priorScreenUpdating = excelApp.ScreenUpdating;
            bool priorEnableEvents = excelApp.EnableEvents;
            Excel.XlCalculation priorCalculation = excelApp.Calculation;
            bool settingsChanged = false;

            try
            {
                // Collect both columns
                Excel.Range columnRange1 = QTUtils.ColumnSelection(
                    excelApp,
                    allowMultipleColumns: false,
                    forcePrompt: true,
                    prompt: "Select the FIRST column:\n(any open workbook)",
                    title: "Range Selector");

                if (columnRange1 == null)
                    return;

                QTUtils.ActivateRange(columnRange1);

                Excel.Range columnRange2 = QTUtils.ColumnSelection(
                    excelApp,
                    allowMultipleColumns: false,
                    forcePrompt: true,
                    prompt: "Select the SECOND column:\n(any open workbook)",
                    title: "Range Selector");

                if (columnRange2 == null)
                    return;

                Excel.Worksheet sheet1 = columnRange1.Worksheet;
                Excel.Worksheet sheet2 = columnRange2.Worksheet;

                bool sameSheet = QTUtils.IsSameSheet(sheet1, sheet2);

                int col1 = columnRange1.Column;
                int col2 = columnRange2.Column;

                if (sameSheet && col1 == col2)
                {
                    MessageBox.Show(
                        "Please select two different columns.",
                        "Compare Columns",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (col1 >= sheet1.Columns.Count || col2 >= sheet2.Columns.Count)
                {
                    MessageBox.Show(
                        "There is no room to insert a column to the right of the selection.",
                        "Compare Columns",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Check if data exists before inserting anything
                int lastRow1 = QTUtils.LastDataRowOrZero(sheet1, col1, FirstDataRow);
                int lastRow2 = QTUtils.LastDataRowOrZero(sheet2, col2, FirstDataRow);

                bool empty1 = lastRow1 < FirstDataRow;
                bool empty2 = lastRow2 < FirstDataRow;

                if (empty1 || empty2)
                {
                    string which;
                    if (empty1 && empty2)
                        which = "Both selected columns are empty.";
                    else if (empty1)
                        which = "Column " + QTUtils.GetColumnLetter(col1) +
                                " on " + sheet1.Name + " is empty.";
                    else
                        which = "Column " + QTUtils.GetColumnLetter(col2) +
                                " on " + sheet2.Name + " is empty.";

                    MessageBox.Show(
                        which + "\nThere is nothing to compare.",
                        "Compare Columns",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                settingsChanged = true;
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;
                excelApp.DisplayStatusBar = true;
                excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                // Insert the two columns
                // On a shared sheet re-adust the second column
                int helper1 = col1 + 1;
                ((Excel.Range)sheet1.Columns[helper1]).Insert(
                    Excel.XlInsertShiftDirection.xlShiftToRight,
                    Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                if (sameSheet && col2 >= helper1)
                    col2++;

                int helper2 = col2 + 1;
                ((Excel.Range)sheet2.Columns[helper2]).Insert(
                    Excel.XlInsertShiftDirection.xlShiftToRight,
                    Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                if (sameSheet && col1 >= helper2)
                {
                    col1++;
                    helper1++;
                }

                // Read both columns
                object[] values1 = ReadColumnChunked(excelApp, sheet1, col1, lastRow1);
                object[] values2 = ReadColumnChunked(excelApp, sheet2, col2, lastRow2);

                // Index
                HashSet<string> keys1 = BuildKeySet(values1);
                HashSet<string> keys2 = BuildKeySet(values2);

                // Write results
                WriteExistsColumn(excelApp, sheet1, helper1, sheet2, col2,
                    values1, keys2, lastRow1, sameSheet);
                WriteExistsColumn(excelApp, sheet2, helper2, sheet1, col1,
                    values2, keys1, lastRow2, sameSheet);

                QTUtils.ActivateRange((Excel.Range)sheet1.Cells[HeaderRow, helper1]);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Compare Columns could not complete:\n\n" + ex.Message,
                    "Compare Columns",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (settingsChanged)
                {
                    excelApp.Calculation = priorCalculation;
                    excelApp.EnableEvents = priorEnableEvents;
                    excelApp.ScreenUpdating = priorScreenUpdating;
                    excelApp.StatusBar = false;
                }
            }
        }

        // Reads a column into a flat 0-based array, one chunk per COM call
        private static object[] ReadColumnChunked(
            Excel.Application excelApp,
            Excel.Worksheet sheet,
            int column,
            int lastRow)
        {
            int totalRows = lastRow - FirstDataRow + 1;
            if (totalRows <= 0)
                return new object[0];

            object[] values = new object[totalRows];

            for (int startRow = FirstDataRow; startRow <= lastRow; startRow += CHUNK_SIZE)
            {
                int endRow = Math.Min(startRow + CHUNK_SIZE - 1, lastRow);
                int rowsToProcess = endRow - startRow + 1;

                Excel.Range source = sheet.Range[
                    sheet.Cells[startRow, column],
                    sheet.Cells[endRow, column]];

                object raw = source.Value2;

                // A single cell returns a scalar rather than a 2D array.
                object[,] block = raw as object[,];

                if (block == null)
                {
                    values[startRow - FirstDataRow] = raw;
                }
                else
                {
                    for (int i = 0; i < rowsToProcess; i++)
                        values[startRow - FirstDataRow + i] = block[i + 1, 1];
                }

            }

            return values;
        }

        // Build a set of normalized keys for fast existence checking
        private static HashSet<string> BuildKeySet(object[] values)
        {
            HashSet<string> set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < values.Length; i++)
            {
                string key = NormalizeKey(values[i]);
                if (key != null)
                    set.Add(key);
            }

            return set;
        }

        // Converts a cell value to a comparable key
        private static string NormalizeKey(object value)
        {
            if (value == null)
                return null;

            if (value is double)
            {
                // "R" round-trips, so 1 and 1.0 produce the same key
                return ((double)value).ToString("R", CultureInfo.InvariantCulture);
            }

            if (value is bool)
                return ((bool)value) ? "TRUE" : "FALSE";

            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (text == null)
                return null;

            text = text.Trim();
            if (text.Length == 0)
                return null;

            if (TreatNumericTextAsNumber)
            {
                double parsed;
                if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
                    return parsed.ToString("R", CultureInfo.InvariantCulture);
            }

            return text;
        }

        // Fills a column with TRUE/FALSE, chunked, and labels the header
        private static void WriteExistsColumn(
            Excel.Application excelApp,
            Excel.Worksheet sheet,
            int helperColumn,
            Excel.Worksheet otherSheet,
            int otherColumn,
            object[] values,
            HashSet<string> otherKeys,
            int lastRow,
            bool sameSheet)
        {
            string otherLetter = QTUtils.GetColumnLetter(otherColumn);

            Excel.Range header = (Excel.Range)sheet.Cells[HeaderRow, helperColumn];
            header.Value2 = sameSheet
                ? "Exists in Column " + otherLetter + "?"
                : "Exists in " + otherSheet.Name + " Column " + otherLetter + "?";
            //header.Font.Bold = true;

            int totalRows = values.Length;
            if (totalRows <= 0)
                return;

            // Inserted columns inherit formatting from the left so set to general
            Excel.Range fullTarget = sheet.Range[
                sheet.Cells[FirstDataRow, helperColumn],
                sheet.Cells[lastRow, helperColumn]];
            fullTarget.NumberFormat = "General";

            for (int startRow = FirstDataRow; startRow <= lastRow; startRow += CHUNK_SIZE)
            {
                int endRow = Math.Min(startRow + CHUNK_SIZE - 1, lastRow);
                int rowsToProcess = endRow - startRow + 1;

                object[,] processArray = new object[rowsToProcess, 1];

                for (int i = 0; i < rowsToProcess; i++)
                {
                    string key = NormalizeKey(values[startRow - FirstDataRow + i]);
                    processArray[i, 0] = key != null && otherKeys.Contains(key);
                }

                Excel.Range rangeToUpdate = sheet.Range[
                    sheet.Cells[startRow, helperColumn],
                    sheet.Cells[endRow, helperColumn]];

                rangeToUpdate.Value2 = processArray;
            }
        }
    }
}
