﻿using System;
using System.Runtime.InteropServices;
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
                // Determine the last used row within the selected range
                Excel.Range usedRange = activeSheet.UsedRange;
                int lastUsedRow = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                // Restrict the range to the selected columns and last used row
                return activeSheet.Range[
                    selectedRange.Cells[1, 1],
                    selectedRange.Cells[lastUsedRow, selectedRange.Columns.Count]
                ];
            }

            // For all other cases, return the selected range
            return selectedRange;
        }

        // Set the range values to an array
        public static object[,] GetRangeValues(Excel.Range range)
        {
            if (range.Rows.Count == 1 && range.Columns.Count == 1)
            {
                var values = new object[1, 1];
                values[0, 0] = range.Value2;
                return values;
            }
            return range.Value2 as object[,];
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
        public static void CleanupResources(Excel.Range range, Excel.Application excelApp)
        {
            if (range != null && Marshal.IsComObject(range))
            {
                Marshal.ReleaseComObject(range);
            }
            excelApp.ScreenUpdating = true;
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
        public static Excel.Range ColumnSelection(Excel.Application excelApp, bool allowMultipleColumns = false)
        {
            if (excelApp == null)
                throw new ArgumentNullException(nameof(excelApp));

            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            if (activeWorkbook == null)
                return null;

            Excel.Worksheet activeSheet = excelApp.ActiveSheet;
            if (activeSheet == null)
                return null;

            Excel.Range selectedRange = excelApp.Selection;

            // Validate initial selection
            bool isValidSelection = ValidateColumnSelection(selectedRange, activeSheet, allowMultipleColumns);

            if (!isValidSelection)
            {
                // Prompt the user to select column(s) using InputBox
                object rangeInput = excelApp.InputBox(
                    allowMultipleColumns
                        ? "You must select one or more columns first:"
                        : "You must select a column first:",
                    "Range Selector",
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    8 // Type 8 allows range selection
                );

                // Handle InputBox cancellation
                if (rangeInput is bool && (bool)rangeInput == false)
                {
                    return null;
                }

                selectedRange = rangeInput as Excel.Range;

                // Validate the user's new selection
                if (!ValidateColumnSelection(selectedRange, activeSheet, allowMultipleColumns))
                {
                    // Invalid selection, treat as canceled
                    return null;
                }
                else
                {
                    selectedRange.Select();
                }
            }

            // Return the valid column selection
            return selectedRange;
        }

        // Validate column selection
        public static bool ValidateColumnSelection(Excel.Range range, Excel.Worksheet sheet, bool allowMultipleColumns)
        {
            if (range == null)
                return false;

            int selectedColumnCount = range.Columns.Count;
            int sheetRowCount = sheet.Rows.Count;

            if (allowMultipleColumns)
            {
                // Check if multiple columns span the entire row range
                return range.Rows.Count == sheetRowCount && selectedColumnCount >= 1;
            }
            else
            {
                // Check if a single column spans the entire row range
                return range.Rows.Count == sheetRowCount && selectedColumnCount == 1;
            }
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
    }
}