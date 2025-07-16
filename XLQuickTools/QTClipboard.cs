using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    internal sealed class QTClipboard : IDisposable
    {
        private static readonly object _lock = new object();
        private static volatile QTClipboard _instance;
        private bool _disposed;

        public static QTClipboard Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new QTClipboard();
                        }
                    }
                }
                return _instance;
            }
        }

        private object[,] _originalValues;
        private string _originalNumberFormat;
        private string _lastFormattedRangeAddress;
        private string _lastWorksheetName;
        private string _lastWorkbookName;

        private QTClipboard()
        {
            _disposed = false;
        }

        public void CopyAndStoreFormat(Excel.Range selectedRange)
        {
            if (selectedRange == null)
            {
                throw new ArgumentNullException(nameof(selectedRange), "Selected range cannot be null.");
            }

            try
            {
                // Store workbook, worksheet and range information
                _lastWorkbookName = selectedRange.Worksheet.Parent.Name;
                _lastWorksheetName = selectedRange.Worksheet.Name;
                _lastFormattedRangeAddress = selectedRange.Address[false, false];

                // Store the format
                _originalNumberFormat = selectedRange.NumberFormat as string ?? string.Empty;

                // Handle single cell vs multi-cell differently
                if (selectedRange.Rows.Count == 1 && selectedRange.Columns.Count == 1)
                {
                    // Single cell: Convert to a 2D array
                    _originalValues = new object[1, 1];
                    _originalValues[0, 0] = selectedRange.Value2;
                }
                else
                {
                    // Multi-cell: Direct cast
                    _originalValues = selectedRange.Value2 as object[,];

                    // Use flat index to get 2nd cell
                    if (selectedRange.Cells.Count >= 2)
                    {
                        Excel.Range secondCell = selectedRange.Cells[2] as Excel.Range;
                        // Use 2nd cell number format instead
                        _originalNumberFormat = secondCell.NumberFormat as string ?? string.Empty;
                    }
                }

                // Update UI
                Globals.Ribbons.Ribbon1.BtnUndo.Enabled = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in CopyAndStoreFormat: {ex.Message}");
                throw;
            }
        }

        public void UndoFormatting()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(QTClipboard));
            }

            if (string.IsNullOrEmpty(_lastFormattedRangeAddress) || _originalValues == null)
            {
                return;
            }

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                excelApp.ScreenUpdating = false;

                // Get the workbook
                workbook = excelApp.Workbooks[_lastWorkbookName];
                if (workbook == null) return;

                // Get the worksheet
                worksheet = workbook.Worksheets[_lastWorksheetName];
                if (worksheet == null) return;

                // Get the range
                range = worksheet.Range[_lastFormattedRangeAddress];
                if (range == null) return;

                // Restore original values and format
                range.Value2 = _originalValues;
                if (!string.IsNullOrEmpty(_originalNumberFormat))
                {
                    range.NumberFormat = _originalNumberFormat;
                }
                else
                {
                    range.NumberFormat = "General";
                }

                    // Update UI
                    Globals.Ribbons.Ribbon1.BtnUndo.Enabled = false;
                range.Select();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error during undo operation: {ex.Message}");
                throw;
            }
            finally
            {
                excelApp.ScreenUpdating = true;

                // Clean up COM objects
                if (range != null && Marshal.IsComObject(range))
                    Marshal.ReleaseComObject(range);
                if (worksheet != null && Marshal.IsComObject(worksheet))
                    Marshal.ReleaseComObject(worksheet);
                if (workbook != null && Marshal.IsComObject(workbook))
                    Marshal.ReleaseComObject(workbook);
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _originalValues = null;
                _originalNumberFormat = null;
                _lastFormattedRangeAddress = null;
                _lastWorksheetName = null;
                _lastWorkbookName = null;
                _disposed = true;
                GC.SuppressFinalize(this);
            }
        }
    }
}