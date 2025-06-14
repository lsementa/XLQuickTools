using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    public partial class HighlightForm : Form
    {
        private readonly Excel.Application _excelApp;

        public HighlightForm(Excel.Application excelApp)
        {
            InitializeComponent();
            _excelApp = excelApp;
        }

        // On Load
        private void HighlightForm_Load(object sender, EventArgs e)
        {
            Excel.Workbook workbook = _excelApp.ActiveWorkbook;

            CbContainsHeaders.Checked = true;
            TbSheetName.Text = QTUtils.GetUniqueName("Highlighted Rows", workbook);
            TbSheetName.Select();
        }

        // Ok button
        private void HighlightForm_Ok_Click(object sender, EventArgs e)
        {
            // Copy highlighted rows
            CopyHighlightedRowsToNewSheet();
            this.Close();
        }

        private static bool HasAnyCellHighlighted(Excel.Range range)
        {
            // Iterate through each cell
            foreach (Excel.Range cell in range.Cells)
            {
                // Check if the cell's interior color is not the default 'no fill' color.
                if (cell.Interior.ColorIndex != (int)Excel.XlColorIndex.xlColorIndexNone)
                {
                    return true;
                }
            }
            return false;
        }

        // Copy highlighted rows to new sheet
        private void CopyHighlightedRowsToNewSheet()
        {
            Excel.Workbook activeWorkbook = _excelApp.ActiveWorkbook;
            Excel.Worksheet activeSheet = _excelApp.ActiveSheet;
            Excel.Range usedRange = activeSheet.UsedRange;

            string sheetName = TbSheetName.Text.Trim();

            if (usedRange == null || (usedRange.Cells.Count == 1 && usedRange.Value2 == null))
            {
                return;
            }

            try
            {
                _excelApp.ScreenUpdating = false;
                Excel.Worksheet newSheet = QTUtils.AddUniqueNamedWorksheet(activeWorkbook, activeSheet, sheetName);

                int newRowIndex = 1;
                int startRowIndex = 1;

                // First row contains headers
                if (CbContainsHeaders.Checked)
                {
                    // Copy the header row to the first row of the new sheet
                    Excel.Range headerRow = (Excel.Range)usedRange.Rows[1];
                    headerRow.Copy(newSheet.Cells[newRowIndex, 1]);
                    newRowIndex++;
                    startRowIndex = 2;
                }

                int totalRows = usedRange.Rows.Count;

                // Iterate through the data rows (skipping header if needed)
                for (int i = startRowIndex; i <= totalRows; i++)
                {
                    Excel.Range row = (Excel.Range)usedRange.Rows[i];
                    if (HasAnyCellHighlighted(row))
                    {
                        row.Copy(newSheet.Cells[newRowIndex, 1]);
                        newRowIndex++;
                    }
                }

                // Copy column widths
                Excel.Range newUsedRange = newSheet.UsedRange;
                QTFormat.CopyColumnWidths(usedRange, newUsedRange);

            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
            finally
            {
                _excelApp.ScreenUpdating = true;
                QTUtils.CleanupResources(usedRange);
            }
        }

        // Cancel button
        private void HighlightForm_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
