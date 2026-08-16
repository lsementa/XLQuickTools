using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    internal class QTUtils
    {
        // Options for removing rows or columns
        public enum DeleteOption
        {
            Rows,
            Columns
        }

        // Method to get the selected range to process
        public static Excel.Range GetRangeToProcess(Excel.Application excelApp)
        {
            if (excelApp == null) return null;

            // Get the active sheet
            Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;
            if (activeSheet == null) return null;

            // Get the selected range
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;
            if (selectedRange == null) return null;

            // Check if the entire worksheet is selected
            if (selectedRange.Address == activeSheet.Cells.Address)
            {
                return activeSheet.UsedRange;
            }

            // Check if one or more entire columns are selected
            bool isEntireColumnsSelected =
                selectedRange.Columns.Count > 0 &&
                selectedRange.Rows.Count == activeSheet.Rows.Count;

            if (isEntireColumnsSelected)
            {
                // UsedRange is filter-agnostic
                Excel.Range used = activeSheet.UsedRange;
                int lastUsedRow = used == null ? 1 : used.Row + used.Rows.Count - 1;
                if (lastUsedRow < 1) lastUsedRow = 1;

                // Restrict the range to the selected columns and last used row
                return activeSheet.Range[
                    selectedRange.Cells[1, 1],
                    activeSheet.Cells[lastUsedRow, selectedRange.Columns[selectedRange.Columns.Count].Column]
                ];
            }

            // For all other cases, return the selected range
            return selectedRange;
        }

        // Check if the sheet has any active filters hiding rows
        private static bool HasActiveFilter(Excel.Worksheet sheet)
        {
            if (sheet.AutoFilterMode && sheet.AutoFilter != null)
            {
                foreach (Excel.Filter filter in sheet.AutoFilter.Filters)
                {
                    if (filter.On) return true;
                }
            }
            return false;
        }

        // Set the range values to an array (filter-aware)
        public static object[,] GetRangeValues(Excel.Range range)
        {
            Excel.Worksheet sheet = range.Worksheet;

            if (!HasActiveFilter(sheet))
            {
                return GetRangeValuesUnfiltered(range);
            }

            Excel.Range visibleCells;
            try
            {
                visibleCells = range.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
            }
            catch
            {
                return new object[0, 0];
            }

            int colCount = range.Columns.Count;
            int firstCol = range.Column;

            // Pass 1: count visible rows
            int visibleRowCount = 0;
            foreach (Excel.Range area in visibleCells.Areas)
                visibleRowCount += area.Rows.Count;

            if (visibleRowCount == 0) return new object[0, 0];

            var result = new object[visibleRowCount, colCount];
            int destRow = 0;

            // Pass 2: bulk-read each contiguous area
            foreach (Excel.Range area in visibleCells.Areas)
            {
                int areaRowCount = area.Rows.Count;

                Excel.Range fullArea = sheet.Range[
                    sheet.Cells[area.Row, firstCol],
                    sheet.Cells[area.Row + areaRowCount - 1, firstCol + colCount - 1]
                ];

                // Normalize to object[,] regardless of area shape
                object[,] areaValues = NormalizeToGrid(fullArea, areaRowCount, colCount);

                for (int r = 0; r < areaRowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        result[destRow + r, c] = areaValues[r, c];
                    }
                }
                destRow += areaRowCount;
            }

            return result;
        }

        // Handles all edge cases: single cell, single row, single column, or multi-cell
        private static object[,] NormalizeToGrid(Excel.Range range, int rowCount, int colCount)
        {
            var grid = new object[rowCount, colCount];

            // Single cell — Value2 is a scalar
            if (rowCount == 1 && colCount == 1)
            {
                grid[0, 0] = range.Value2;
                return grid;
            }

            // Multi-cell — Value2 is object[,] but 1-based from Excel
            if (rowCount > 1 && colCount > 1)
            {
                object[,] raw = range.Value2 as object[,];
                if (raw == null) return grid;
                for (int r = 0; r < rowCount; r++)
                    for (int c = 0; c < colCount; c++)
                        grid[r, c] = raw[r + 1, c + 1]; // Excel returns 1-based arrays
                return grid;
            }

            // Single row or single column — Value2 comes back as object[,] but
            // still 1-based; handle both orientations explicitly
            object[,] rawEdge = range.Value2 as object[,];
            if (rawEdge != null)
            {
                for (int r = 0; r < rowCount; r++)
                    for (int c = 0; c < colCount; c++)
                        grid[r, c] = rawEdge[r + 1, c + 1];
                return grid;
            }

            // Fallback: scalar leaked through (shouldn't happen)
            grid[0, 0] = range.Value2;
            return grid;
        }

        // Reads a range into a 1-based object[,] regardless of its shape.
        // Range.Value/Value2 hands back a scalar for a single cell, which makes any
        // direct cast to object[,] fail; this always returns a 1-based grid.
        public static object[,] GetValueArray(Excel.Range range)
        {
            if (range == null) return null;

            object raw = range.Value2;

            object[,] arr = raw as object[,];
            if (arr != null) return arr;

            // Single cell (or empty) - 1x1 array with lower bounds of 1
            object[,] single = (object[,])Array.CreateInstance(
                typeof(object), new[] { 1, 1 }, new[] { 1, 1 });
            single[1, 1] = raw;
            return single;
        }

        // Original bulk read
        private static object[,] GetRangeValuesUnfiltered(Excel.Range range)
        {
            if (range.Rows.Count == 1 && range.Columns.Count == 1)
            {
                var values = new object[1, 1];
                values[0, 0] = range.Value2;
                return values;
            }

            // Normalize 1-based Excel array to 0-based
            object[,] raw = range.Value2 as object[,];
            if (raw == null) return new object[0, 0];

            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            var result = new object[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    result[r, c] = raw[r + 1, c + 1];

            return result;
        }

        // Returns values AND the original Excel row number for each row (filter-aware)
        public static (object[,] Values, int[] SourceRows) GetRangeValuesWithRows(Excel.Range range)
        {
            Excel.Worksheet sheet = range.Worksheet;

            if (!HasActiveFilter(sheet))
            {
                // No filter — source rows are sequential from range.Row
                object[,] vals = GetRangeValuesUnfiltered(range);
                int rowCount = vals.GetLength(0);
                int[] rows = new int[rowCount];
                for (int i = 0; i < rowCount; i++)
                    rows[i] = range.Row + i;
                return (vals, rows);
            }

            Excel.Range visibleCells;
            try
            {
                visibleCells = range.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
            }
            catch
            {
                return (new object[0, 0], Array.Empty<int>());
            }

            int colCount = range.Columns.Count;
            int firstCol = range.Column;

            // Pass 1: record Excel row number of every visible row
            var sourceRowList = new List<int>();
            foreach (Excel.Range area in visibleCells.Areas)
                for (int r = 0; r < area.Rows.Count; r++)
                    sourceRowList.Add(area.Row + r);

            int visibleRowCount = sourceRowList.Count;
            if (visibleRowCount == 0) return (new object[0, 0], Array.Empty<int>());

            var result = new object[visibleRowCount, colCount];
            int destRow = 0;

            // Pass 2: bulk-read each contiguous area
            foreach (Excel.Range area in visibleCells.Areas)
            {
                int areaRowCount = area.Rows.Count;

                Excel.Range fullArea = sheet.Range[
                    sheet.Cells[area.Row, firstCol],
                    sheet.Cells[area.Row + areaRowCount - 1, firstCol + colCount - 1]
                ];

                object[,] areaValues = NormalizeToGrid(fullArea, areaRowCount, colCount);

                for (int r = 0; r < areaRowCount; r++)
                    for (int c = 0; c < colCount; c++)
                        result[destRow + r, c] = areaValues[r, c];

                destRow += areaRowCount;
            }

            return (result, sourceRowList.ToArray());
        }

        // Write processed values back to their exact original row positions (filter-aware)
        public static void SetRangeValues(Excel.Range originalRange, object[,] values, int[] sourceRows)
        {
            Excel.Worksheet sheet = originalRange.Worksheet;
            int firstCol = originalRange.Column;
            int colCount = originalRange.Columns.Count;
            int rowCount = sourceRows.Length;

            int i = 0;
            while (i < rowCount)
            {
                // Find the end of this consecutive run of row numbers
                int j = i;
                while (j + 1 < rowCount && sourceRows[j + 1] == sourceRows[j] + 1)
                    j++;

                int areaRowCount = j - i + 1;

                // Build sub-array for this contiguous block
                var areaValues = new object[areaRowCount, colCount];
                for (int r = 0; r < areaRowCount; r++)
                    for (int c = 0; c < colCount; c++)
                        areaValues[r, c] = values[i + r, c];

                // Write the whole block in one COM call
                Excel.Range writeArea = sheet.Range[
                    sheet.Cells[sourceRows[i], firstCol],
                    sheet.Cells[sourceRows[j], firstCol + colCount - 1]
                ];
                writeArea.Value2 = areaValues;

                i = j + 1;
            }
        }

        // Method tp delete empty rows
        public static int DeleteRows(Excel.Worksheet sheet, Excel.Application app, int lastRow, int startRow)
        {
            int rowsDeleted = 0;

            // Start from the bottom and work upwards
            for (int row = lastRow; row >= startRow; row--)
            {
                Excel.Range rowRange = sheet.Rows[row];
                if (app.WorksheetFunction.CountA(rowRange) == 0)
                {
                    rowRange.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    rowsDeleted++;
                }
            }

            return rowsDeleted;
        }

        // Method to delete empty columns
        public static int DeleteColumns(Excel.Worksheet sheet, Excel.Application app, int lastColumn, int startColumn, int lastRow)
        {
            int colsDeleted = 0;

            // Start from the rightmost column and work left
            for (int col = lastColumn; col >= startColumn; col--)
            {
                // Column range including Row 1
                Excel.Range colRange = sheet.Range[sheet.Cells[1, col], sheet.Cells[lastRow, col]];

                // Column range ignoring Row 1 for CountA
                if (app.WorksheetFunction.CountA(colRange.Offset[1, 0]) == 0)
                {
                    colRange.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
                    colsDeleted++;
                }
            }

            return colsDeleted;
        }

        // Check if worksheet exists method
        public static bool WorksheetExists(Excel.Workbook workbook, string sheetName)
        {
            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    return true;
                }
            }
            return false;
        }

        // Get unique name
        public static string GetUniqueName(string baseName, Excel.Workbook workbook)
        {
            string uniqueName = baseName;
            int counter = 1;

            while (WorksheetExists(workbook, uniqueName))
            {
                uniqueName = $"{baseName}{counter}";
                counter++;
            }

            return uniqueName;
        }

        // Create unique worksheet
        public static Excel.Worksheet AddUniqueNamedWorksheet(Excel.Workbook workbook, Excel.Worksheet worksheet, string baseName)
        {
            Excel.Worksheet newSheet = (Excel.Worksheet)workbook.Sheets.Add(After: worksheet);
            newSheet.Name = GetUniqueName(baseName, workbook);
            return newSheet;
        }

        // Convert column index to letter (A, B, C, ... AA, AB, etc.)
        public static string GetColumnLetter(int colIndex)
        {
            string columnLetter = "";
            while (colIndex > 0)
            {
                colIndex--;
                columnLetter = (char)('A' + (colIndex % 26)) + columnLetter;
                colIndex /= 26;
            }
            return columnLetter;
        }

        // Show error message
        public static void ShowError(Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                $"An error occurred: {ex.Message}\nStack trace: {ex.StackTrace}",
                "Error",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }

        // Method to clean up resources
        public static void CleanupResources(Excel.Range range)
        {
            if (range != null && Marshal.IsComObject(range))
            {
                Marshal.ReleaseComObject(range);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // Method to return the delimiter
        public static string GetDelimiter(string delimText, string customValue)
        {
            String delimiter = null;

            // Determine what delimiter to use
            switch (delimText)
            {
                case "Space":
                    delimiter = " ";
                    break;
                case "Tab":
                    delimiter = "\t";
                    break;
                case "Carriage Return":
                    delimiter = "\r";
                    break;
                case "Line Feed (Newline)":
                    delimiter = "\n";
                    break;
                case "Vertical Tab":
                    delimiter = "\v";
                    break;
                case "Form Feed":
                    delimiter = "\f";
                    break;
                case "Carriage Return and Line Feed":
                    delimiter = "\r\n";
                    break;
                case "Non-breaking Space":
                    delimiter = "\u00A0"; // Unicode for non-breaking space
                    break;
                case "--Custom--":
                    delimiter = customValue.Trim();
                    break;
            }
            return delimiter;
        }

        // Returns the corresponding Excel NumberFormat string
        public static string GetExcelNumberFormat(string format)
        {
            switch (format)
            {
                // Date Formats
                case "yyyy-MM-dd":
                    return "yyyy-mm-dd;@";
                case "dd/MM/yyyy":
                    return "dd/mm/yyyy;@";
                case "d/M/yyyy":
                    return "d/m/yyyy;@";
                case "dd MMM yyyy":
                    return "dd mmm yyyy;@";
                case "dd MMMM yyyy":
                    return "dd mmmm yyyy;@";
                case "yyyy/MM/dd":
                    return "yyyy/mm/dd;@";
                case "yyyy.MM.dd":
                    return "yyyy.mm.dd;@";
                case "yyyy MMM dd":
                    return "yyyy mmm dd;@";
                case "M/d/yyyy":
                    return "m/d/yyyy;@";
                case "MM/dd/yyyy":
                    return "mm/dd/yyyy;@";
                case "MMM dd, yyyy":
                    return "mmm dd, yyyy;@";
                case "MMMM dd, yyyy":
                    return "mmmm dd, yyyy;@";

                // ZIP Code Formats
                case "#####-####":
                    return "00000-0000";

                // Phone Number Formats
                case "(###) ###-####":
                    return "[<=9999999]###-####;(###) ###-####";

                // SSN
                case "###-##-####":
                    return "000-00-0000";

                // Default to text
                default:
                    return "@";
            }
        }

        // Prompt for column selection
        public static Excel.Range ColumnSelection(Excel.Application excelApp,
            bool allowMultipleColumns = false,
                bool forcePrompt = false, string prompt = null,
                string title = "Range Selector")
        {
            if (excelApp == null)
                throw new ArgumentNullException(nameof(excelApp));

            if (excelApp.ActiveWorkbook == null)
                return null;

            // Accept whatever is already selected, unless we were told not to
            if (!forcePrompt)
            {
                Excel.Range current = excelApp.Selection as Excel.Range;
                if (IsFullColumnSelection(current, allowMultipleColumns))
                    return current;
            }

            if (string.IsNullOrEmpty(prompt))
            {
                prompt = allowMultipleColumns
                    ? "Select one or more entire columns:"
                    : "Select an entire column:";
            }

            object rangeInput;
            try
            {
                // Type 8 = range selection. The user may switch workbooks/sheets
                // while this is open; the returned Range carries its own parents.
                rangeInput = excelApp.InputBox(
                    prompt,
                    title,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    8);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Esc, or text that won't resolve to a range.
                return null;
            }

            // Cancel returns false.
            if (rangeInput is bool)
                return null;

            Excel.Range selectedRange = rangeInput as Excel.Range;

            if (!IsFullColumnSelection(selectedRange, allowMultipleColumns))
            {
                MessageBox.Show(
                    allowMultipleColumns
                        ? "Please select one or more entire columns (click the column headers)."
                        : "Please select an entire column (click the column header).",
                    "Invalid Selection",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return null;
            }

            return selectedRange;
        }

        // Check if the range is a full column selection
        public static bool IsFullColumnSelection(Excel.Range range, bool allowMultipleColumns)
        {
            if (range == null)
                return false;

            try
            {
                Excel.Worksheet sheet = range.Worksheet;
                if (sheet == null)
                    return false;

                if (range.Areas.Count > 1)
                    return false;

                // Entire column => row count matches the sheet's row count.
                if (range.Rows.Count != sheet.Rows.Count)
                    return false;

                if (!allowMultipleColumns && range.Columns.Count != 1)
                    return false;

                return true;
            }
            catch
            {
                // Chart sheet, dead RCW, etc.
                return false;
            }
        }

        // Range.Select() throws unless the owning book and sheet are active
        public static void ActivateRange(Excel.Range range)
        {
            if (range == null)
                return;

            try
            {
                Excel.Worksheet sheet = range.Worksheet;
                Excel.Workbook book = sheet.Parent as Excel.Workbook;

                if (book != null)
                    book.Activate();

                sheet.Activate();
                range.Select();
            }
            catch
            {
                // Cosmetic only.
            }
        }

        //  Check if two worksheets are the same by comparing their parent workbook and name
        public static bool IsSameSheet(Excel.Worksheet a, Excel.Worksheet b)
        {
            if (a == null || b == null)
                return false;

            Excel.Workbook bookA = a.Parent as Excel.Workbook;
            Excel.Workbook bookB = b.Parent as Excel.Workbook;

            if (bookA == null || bookB == null)
                return false;

            return string.Equals(bookA.FullName, bookB.FullName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
        }

        // Last used row in a column, never below firstDataRow
        public static int LastDataRow(Excel.Worksheet sheet, int columnIndex, int firstDataRow)
        {
            Excel.Range bottom = (Excel.Range)sheet.Cells[sheet.Rows.Count, columnIndex];
            int lastRow = bottom.End[Excel.XlDirection.xlUp].Row;
            return lastRow < firstDataRow ? firstDataRow : lastRow;
        }

        // Method to get Effective used range
        public static Excel.Range GetEffectiveUsedRange(Excel.Worksheet sheet)
        {
            var usedRange = sheet.UsedRange;

            // Check if used range is effectively empty
            if (usedRange == null ||
                (usedRange.Cells.Count == 1 &&
                 string.IsNullOrEmpty(usedRange.Cells[1, 1].Value?.ToString())))
            {
                // If used range is empty, return entire sheet
                return sheet.Cells;
            }

            return usedRange;
        }

        // Select folder location dialog
        public static string SelectSaveFolder()
        {
            // Use FolderBrowserDialog to select a folder
            using (System.Windows.Forms.FolderBrowserDialog folderDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                folderDialog.Description = "Select a folder to save the file";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return folderDialog.SelectedPath;
                }
                else
                {
                    return null;  // No folder selected
                }
            }
        }

        // Convert tables to ranges
        public static void ConvertTablesToRanges()
        {
            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;

            // List of tables
            var tables = new List<Excel.ListObject>();
            foreach (Excel.ListObject table in activeSheet.ListObjects)
            {
                tables.Add(table);
            }

            // Unlist/Convert to range each table
            foreach (Excel.ListObject table in tables)
            {
                table.Unlist();
            }
        }

        // Dictionary to store special character replacements
        public static readonly Dictionary<char, string> specialReplacements = new Dictionary<char, string>
        {
            // --- Scandinavian ---
            ['ø'] = "o",
            ['Ø'] = "O",
            ['æ'] = "ae",
            ['Æ'] = "AE",
            ['å'] = "a",
            ['Å'] = "A",

            // --- German ---
            ['ß'] = "ss",
            ['ẞ'] = "SS",

            // --- French Ligatures ---
            ['œ'] = "oe",
            ['Œ'] = "OE",

            // --- Polish ---
            ['ł'] = "l",
            ['Ł'] = "L",

            // --- Croatian/Slovak/Slovenian ---
            ['đ'] = "d",
            ['Đ'] = "D",

            // --- Icelandic/Old English ---
            ['þ'] = "th",
            ['Þ'] = "Th",
            ['ð'] = "d",
            ['Ð'] = "D",

            // --- Turkish ---
            ['ğ'] = "g",
            ['Ğ'] = "G",
            ['ı'] = "i",
            ['İ'] = "I",
            ['ş'] = "s",
            ['Ş'] = "S",

            // --- Czech/Slovak ---
            ['č'] = "c",
            ['Č'] = "C",
            ['ř'] = "r",
            ['Ř'] = "R",
            ['š'] = "s",
            ['Š'] = "S",
            ['ž'] = "z",
            ['Ž'] = "Z",

            // --- Hungarian ---
            ['ő'] = "o",
            ['Ő'] = "O",
            ['ű'] = "u",
            ['Ű'] = "U",

            // --- Romanian ---
            ['ă'] = "a",
            ['Ă'] = "A",
            ['ș'] = "s",
            ['Ș'] = "S",
            ['ț'] = "t",
            ['Ț'] = "T",

            // --- Cyrillic lowercase ---
            ['а'] = "a",
            ['б'] = "b",
            ['в'] = "v",
            ['г'] = "g",
            ['д'] = "d",
            ['е'] = "e",
            ['ё'] = "yo",
            ['ж'] = "zh",
            ['з'] = "z",
            ['и'] = "i",
            ['й'] = "y",
            ['к'] = "k",
            ['л'] = "l",
            ['м'] = "m",
            ['н'] = "n",
            ['о'] = "o",
            ['п'] = "p",
            ['р'] = "r",
            ['с'] = "s",
            ['т'] = "t",
            ['у'] = "u",
            ['ф'] = "f",
            ['х'] = "kh",
            ['ц'] = "ts",
            ['ч'] = "ch",
            ['ш'] = "sh",
            ['щ'] = "shch",
            ['ъ'] = "",
            ['ы'] = "y",
            ['ь'] = "",
            ['э'] = "e",
            ['ю'] = "yu",
            ['я'] = "ya",

            // --- Cyrillic uppercase ---
            ['А'] = "A",
            ['Б'] = "B",
            ['В'] = "V",
            ['Г'] = "G",
            ['Д'] = "D",
            ['Е'] = "E",
            ['Ё'] = "Yo",
            ['Ж'] = "Zh",
            ['З'] = "Z",
            ['И'] = "I",
            ['Й'] = "Y",
            ['К'] = "K",
            ['Л'] = "L",
            ['М'] = "M",
            ['Н'] = "N",
            ['О'] = "O",
            ['П'] = "P",
            ['Р'] = "R",
            ['С'] = "S",
            ['Т'] = "T",
            ['У'] = "U",
            ['Ф'] = "F",
            ['Х'] = "Kh",
            ['Ц'] = "Ts",
            ['Ч'] = "Ch",
            ['Ш'] = "Sh",
            ['Щ'] = "Shch",
            ['Ъ'] = "",
            ['Ы'] = "Y",
            ['Ь'] = "",
            ['Э'] = "E",
            ['Ю'] = "Yu",
            ['Я'] = "Ya",

            // --- Greek lowercase ---
            ['α'] = "a",
            ['β'] = "b",
            ['γ'] = "g",
            ['δ'] = "d",
            ['ε'] = "e",
            ['ζ'] = "z",
            ['η'] = "i",
            ['θ'] = "th",
            ['ι'] = "i",
            ['κ'] = "k",
            ['λ'] = "l",
            ['μ'] = "m",
            ['ν'] = "n",
            ['ξ'] = "x",
            ['ο'] = "o",
            ['π'] = "p",
            ['ρ'] = "r",
            ['σ'] = "s",
            ['ς'] = "s",
            ['τ'] = "t",
            ['υ'] = "y",
            ['φ'] = "f",
            ['χ'] = "ch",
            ['ψ'] = "ps",
            ['ω'] = "o",
            ['ϑ'] = "th",
            ['ϒ'] = "Y",
            ['ϖ'] = "p",

            // --- Greek uppercase ---
            ['Α'] = "A",
            ['Β'] = "B",
            ['Γ'] = "G",
            ['Δ'] = "D",
            ['Ε'] = "E",
            ['Ζ'] = "Z",
            ['Η'] = "I",
            ['Θ'] = "Th",
            ['Ι'] = "I",
            ['Κ'] = "K",
            ['Λ'] = "L",
            ['Μ'] = "M",
            ['Ν'] = "N",
            ['Ξ'] = "X",
            ['Ο'] = "O",
            ['Π'] = "P",
            ['Ρ'] = "R",
            ['Σ'] = "S",
            ['Τ'] = "T",
            ['Υ'] = "Y",
            ['Φ'] = "F",
            ['Χ'] = "Ch",
            ['Ψ'] = "Ps",
            ['Ω'] = "O",

            // --- Arabic-Indic digits ---
            ['٠'] = "0",
            ['١'] = "1",
            ['٢'] = "2",
            ['٣'] = "3",
            ['٤'] = "4",
            ['٥'] = "5",
            ['٦'] = "6",
            ['٧'] = "7",
            ['٨'] = "8",
            ['٩'] = "9",

            // --- Eastern Arabic-Indic digits ---
            ['۰'] = "0",
            ['۱'] = "1",
            ['۲'] = "2",
            ['۳'] = "3",
            ['۴'] = "4",
            ['۵'] = "5",
            ['۶'] = "6",
            ['۷'] = "7",
            ['۸'] = "8",
            ['۹'] = "9",

            // --- Mathematical symbols ---
            ['×'] = "x",
            ['÷'] = "/",
            ['±'] = "+/-",
            ['≠'] = "!=",
            ['≤'] = "<=",
            ['≥'] = ">=",
            ['∞'] = "[infinity]",
            ['√'] = "[sqrt]",
            ['∑'] = "[sum]",
            ['∫'] = "[integral]",
            ['∆'] = "[delta]",
            ['∇'] = "[nabla]",
            ['∼'] = "~",
            ['≈'] = "~~",
            ['≡'] = "===",
            ['∝'] = "[proportional to]",
            ['∴'] = "[therefore]",
            ['∵'] = "[because]",
            ['∂'] = "[partial]",
            ['∀'] = "[for all]",
            ['∃'] = "[there exists]",
            ['∅'] = "[empty set]",
            ['∈'] = "[in]",
            ['∉'] = "[not in]",

            // --- Fullwidth punctuation ---
            ['！'] = "!",
            ['？'] = "?",
            ['（'] = "(",
            ['）'] = ")",
            ['［'] = "[",
            ['］'] = "]",
            ['｛'] = "{",
            ['｝'] = "}",
            ['〈'] = "<",
            ['〉'] = ">",
            ['．'] = ".",
            ['，'] = ",",
            ['；'] = ";",
            ['：'] = ":",
            ['／'] = "/",
            ['＼'] = "\\",
            ['＋'] = "+",
            ['－'] = "-",
            ['＝'] = "=",
            ['＊'] = "*",
            ['＆'] = "&",
            ['＃'] = "#",
            ['％'] = "%",
            ['＠'] = "@",
            ['｜'] = "|",

            // --- Arrows ---
            ['←'] = "<-",
            ['→'] = "->",
            ['↑'] = "^",
            ['↓'] = "v",
            ['↔'] = "<->",
            ['⇐'] = "<=",
            ['⇒'] = "=>",
            ['⇑'] = "^^",
            ['⇓'] = "vv",
            ['↩'] = "<-",
            ['↪'] = "->",

            // --- Quotation marks ---
            ['\u201C'] = "\"",
            ['\u201D'] = "\"",
            ['\u2018'] = "'",
            ['\u2019'] = "'",
            ['\u201E'] = "\"",
            ['\u201A'] = "'",
            ['\u301D'] = "\"",
            ['\u301E'] = "\"",
            ['\u2039'] = "<",
            ['\u203A'] = ">",

            // --- Dashes ---
            ['\u2013'] = "-",
            ['\u2014'] = "--",
            ['\u2015'] = "--",
            ['\u2010'] = "-",
            ['\u2011'] = "-",
            ['\u2012'] = "-",
            ['\u2043'] = "-",
            ['\u207B'] = "-",
            ['\u208B'] = "-",
            ['\u00AD'] = "",

            // --- Dots and ellipses ---
            ['\u00B7'] = ".",
            ['\u2022'] = "*",
            ['\u2023'] = ">",
            ['\u2026'] = "...",
            ['\u22EF'] = "...",
            ['\u22EE'] = "...",
            ['\u22F0'] = "...",
            ['\u22F1'] = "...",

            // --- Spaces ---
            ['\u00A0'] = " ",
            ['\u2007'] = " ",
            ['\u202F'] = " ",
            ['\u2060'] = "",
            ['\u200B'] = "",
            ['\u200C'] = "",
            ['\u200D'] = "",
            ['\uFEFF'] = "",

            // --- Fractions ---
            ['½'] = "1/2",
            ['⅓'] = "1/3",
            ['⅔'] = "2/3",
            ['¼'] = "1/4",
            ['¾'] = "3/4",
            ['⅕'] = "1/5",
            ['⅖'] = "2/5",
            ['⅗'] = "3/5",
            ['⅘'] = "4/5",
            ['⅙'] = "1/6",
            ['⅚'] = "5/6",
            ['⅛'] = "1/8",
            ['⅜'] = "3/8",
            ['⅝'] = "5/8",
            ['⅞'] = "7/8",

            // --- Miscellaneous ---
            ['¬'] = "not",
            ['¦'] = "|",
            ['ª'] = "a",
            ['º'] = "o"
        };

    }
}