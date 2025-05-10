using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using static XLQuickTools.QTSettings;
using static XLQuickTools.QTUtils;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    internal class QTFunctions
    {

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

                    // Turn off screen updating
                    excelApp.ScreenUpdating = false;

                    // Use AutoFilter to filter unique criteria
                    Excel.Range columnRange = selectedRange.Columns[1];
                    int lastRow = activeSheet.Cells[activeSheet.Rows.Count, columnRange.Column].End(Excel.XlDirection.xlUp).Row;
                    Excel.Range dataRange = activeSheet.Range[activeSheet.Cells[1, 1], activeSheet.Cells[lastRow, activeSheet.Columns.Count]];

                    // Get unique values from the column
                    var uniqueValues = new HashSet<object>();
                    for (int row = 2; row <= lastRow; row++) // Skip the header row
                    {
                        object value = columnRange.Cells[row, 1].Value;
                        if (value != null && !uniqueValues.Contains(value))
                        {
                            uniqueValues.Add(value);
                        }
                    }

                    // Check if the number of unique values exceeds the threshold (e.g., 50)
                    int uniqueCount = uniqueValues.Count;
                    if (uniqueCount > 50)
                    {
                        DialogResult result = MessageBox.Show(
                            $"There are {uniqueCount} unique values, which will create a large number of sheets. Do you want to proceed?",
                            "Warning: Large Number of Sheets",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning
                        );

                        if (result == DialogResult.No)
                        {
                            return; // Stop the process if the user clicks "No"
                        }
                    }

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

                        // AutoFit columns
                        newSheet.Columns.AutoFit();

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
                }
            }
        }

        // Find duplicates
        public static void FindDuplicates()
        {
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

                    // Check if the next column is a "Count" column
                    if (adjacentHeaderCell.Value != null && adjacentHeaderCell.Value.ToString() == "Count")
                    {
                        // Remove Autofilter if applied
                        if (activeSheet.AutoFilterMode) activeSheet.AutoFilterMode = false;

                        // Delete the "Count" column
                        adjacentHeaderCell.EntireColumn.Delete();
                        return;
                    }

                    // Get the selected column range
                    Excel.Range columnRange = selectedRange.Columns[1];

                    // Read the values from the selected range into an array
                    object[,] columnValues = columnRange.Value2 as object[,];

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
                        // Ensure no previous AutoFilter exists
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

                    }
                    else
                    {
                        MessageBox.Show("No duplicates found in the selected column.", "No Duplicates",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch
                {
                }
                finally
                {
                    // Turn screen updating back on
                    excelApp.ScreenUpdating = true;
                }
            }
        }

        // Add or remove hyperlinks
        public static void ToggleHyperlinks()
        {
            UserSettings settings = LoadUserSettingsFromXml();
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            string selectedURL = string.Empty;

            // Find the use this hyperlink
            var matchingEntry = settings.HyperlinkEntries
                .FirstOrDefault(entry => entry.Use == true);

            // If it exists set all the fields
            if (matchingEntry != null)
            {
                selectedURL = matchingEntry.URL;
            }

            if (activeWorkbook != null && !string.IsNullOrEmpty(selectedURL))
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

                    // Search for "HYPERLINK" formulas in the selected column
                    Excel.Range firstFound = selectedRange.Find(
                        What: "HYPERLINK",
                        After: selectedRange.Cells[1, 1],
                        LookIn: Excel.XlFindLookIn.xlFormulas,
                        LookAt: Excel.XlLookAt.xlPart,
                        SearchOrder: Excel.XlSearchOrder.xlByColumns,
                        SearchDirection: Excel.XlSearchDirection.xlNext,
                        MatchCase: false);

                    if (firstFound != null)
                    {
                        // HYPERLINK formula found, copy and paste values in the column
                        // Set to Text
                        selectedRange.NumberFormat = "@";
                        selectedRange.Value2 = selectedRange.Value2;
                        // Remove hyperlink formatting: Reset font color and underline
                        selectedRange.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone;
                        selectedRange.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    }
                    else
                    {
                        // No hyperlinks found, process the column
                        Excel.Range lastCell = selectedRange.Cells[selectedRange.Rows.Count, 1].End(Excel.XlDirection.xlUp);
                        int lastRow = lastCell.Row;

                        // Read the column values into an array
                        object[,] values = selectedRange.Value2 as object[,];

                        int totalRows = lastRow;
                        int chunkSize = 5000; // Process in chunks

                        // Process the data in chunks
                        for (int startRow = 2; startRow <= totalRows; startRow += chunkSize)
                        {
                            int endRow = Math.Min(startRow + chunkSize - 1, totalRows);
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

                                // Toggle on: Create a hyperlink
                                if (!string.IsNullOrEmpty(cellValue))
                                {
                                    string hyperlinkURL = selectedURL.Replace("{ID}", cellValue).Replace("{id}", cellValue);
                                    string hyperlinkFormula = $"=HYPERLINK(\"{hyperlinkURL}\", \"{cellValue}\")";

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

        // Get the Unique count
        public static int GetUniqueCount(Excel.Range rangeToProcess)
        {
            try
            {
                if (rangeToProcess == null) return 0;

                var values = QTUtils.GetRangeValues(rangeToProcess);

                // Flatten the values into a single list, excluding null or empty cells
                var allValues = values.Cast<object>()
                                      .Where(v => v != null && !string.IsNullOrWhiteSpace(v.ToString()))
                                      .Select(v => v.ToString())
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

        // Get the unique row count with optional clipboard copy
        public static int GetUniqueRows(Excel.Range rangeToProcess, CheckedListBox clbColumns, bool copyToClipboard = false)
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
                object[,] data = rangeToProcess.Value;
                int rowCount = data.GetLength(0);
                int colCount = data.GetLength(1);

                // Process each row
                for (int row = 1; row <= rowCount; row++)
                {
                    // Build a key string from the checked columns for this row
                    StringBuilder keyBuilder = new StringBuilder();
                    bool rowHasValue = false;

                    foreach (int colIndex in checkedColumnIndices)
                    {
                        // Check if the column index is within the range
                        if (colIndex < colCount)
                        {
                            // Get column value
                            object value = data[row, colIndex + 1]; // +1 because Excel arrays are 1-based
                            string strValue = value?.ToString() ?? string.Empty;

                            // Check if this cell has a value
                            if (!string.IsNullOrWhiteSpace(strValue))
                            {
                                rowHasValue = true;
                            }

                            keyBuilder.Append(strValue);
                            keyBuilder.Append("|"); // Use a separator unlikely to appear in cell values
                        }
                    }

                    // Only add non-blank rows to the unique set
                    if (rowHasValue)
                    {
                        string key = keyBuilder.ToString();

                        // If we're copying to clipboard, store the entire row
                        if (copyToClipboard)
                        {
                            // Only store the row data if we need to copy to clipboard
                            if (!uniqueRowData.ContainsKey(key))
                            {
                                // Extract all cell values for this row
                                List<object> rowValues = new List<object>();
                                for (int col = 1; col <= colCount; col++)
                                {
                                    rowValues.Add(data[row, col]);
                                }

                                uniqueRowData.Add(key, rowValues);
                            }
                        }
                        else
                        {
                            // If not copying, just track the unique keys
                            if (!uniqueRowData.ContainsKey(key))
                            {
                                uniqueRowData.Add(key, null); // Just use null as we don't need the data
                            }
                        }
                    }
                }

                // Copy to clipboard if requested
                if (copyToClipboard && uniqueRowData.Count > 0)
                {
                    // Create a string builder for the clipboard text
                    StringBuilder clipboardContent = new StringBuilder();

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
            Excel.Worksheet activeSheet = excelApp.ActiveSheet;
            Excel.Range rangeToProcess = null;

            try
            {
                rangeToProcess = QTUtils.GetRangeToProcess(excelApp);
                if (rangeToProcess == null) return;

                // Check if the range is a single cell
                if (rangeToProcess.Rows.Count == 1 && rangeToProcess.Columns.Count == 1)
                {
                    // Get the CurrentRegion of the selected cell
                    rangeToProcess = rangeToProcess.CurrentRegion;
                    // Check if valid range
                    if (rangeToProcess.Rows.Count == 1 && rangeToProcess.Columns.Count == 1)
                    {
                        // Get the used range
                        rangeToProcess = activeSheet.UsedRange;
                        // Check if valid range
                        if (rangeToProcess.Rows.Count == 1 && rangeToProcess.Columns.Count == 1)
                            return;
                    }
                }

                // Show form without selecting the range first
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
    }
}
