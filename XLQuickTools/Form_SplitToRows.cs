using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    public partial class SplitterForm : Form
    {
        private readonly Excel.Worksheet _activeSheet;
        private Excel.Range usedRange;
        private Excel.Range originalRange;

        public SplitterForm(Excel.Worksheet activeSheet)
        {
            InitializeComponent();
            _activeSheet = activeSheet;
        }

        // On Load
        private void SplitterForm_Load(object sender, EventArgs e)
        {
            // Process pending Windows messages to clear the spinning cursor
            Application.DoEvents();

            // Get the used range and show the user
            usedRange = _activeSheet.UsedRange;
            usedRange.Select();

            // Check if the range has more than one row and starts at row 1
            if (usedRange.Rows.Count >= 2 && usedRange.Row == 1)
            {
                // Automatically check there are headers
                CbHeaders.Checked = true;
            }
            else
            {
                // Disable checkbox for headers if row count only one
                if (usedRange.Rows.Count == 1)
                {
                    CbHeaders.Enabled = false;
                }
            }

            // Populate the delimiter combobox with options
            this.CbDelimiter.Items.AddRange(new object[]
            {
                "Tab",
                "Space",
                "Carriage Return",
                "Line Feed (Newline)",
                "Vertical Tab",
                "Form Feed",
                "Carriage Return and Line Feed",
                "Non-breaking Space",
                "--Custom--"
            });

            // Set the dropdown to custom
            this.CbDelimiter.SelectedItem = "--Custom--";

            // Put the cursor in the custom textbox
            this.TbCustom.Select();
        }

        // Main method to split delimted columns to new rows
        public void SplitToRows(Excel.Worksheet activeSheet, string delimText, string customValue)
        {
            var excelApp = Globals.ThisAddIn.Application;
            if (usedRange == null) return;

            try
            {
                excelApp.ScreenUpdating = false;

                // Determine actual bounds of usedRange
                int startRow = usedRange.Row;
                int endRow = startRow + usedRange.Rows.Count - 1;
                int startCol = usedRange.Column;
                int endCol = startCol + usedRange.Columns.Count - 1;

                int changeCnt = 0;
                string delimiter = QTUtils.GetDelimiter(delimText, customValue);

                // Start from the bottom row and work your way up
                for (int row = endRow; row >= startRow; row--)
                {
                    int maxCnt = 0;

                    // Column iteration: First pass
                    for (int col = startCol; col <= endCol; col++)
                    {
                        var cell = activeSheet.Cells[row, col] as Excel.Range;
                        string cellValue = null;

                        // Work with the string value
                        if (cell != null && cell.Value2 != null)
                        {
                            cellValue = cell.Value2.ToString();
                        }
                        
                        // Get the max amount of rows needed
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            string[] temp = cellValue.Split(new[] { delimiter }, StringSplitOptions.None);
                            int cntDelim = temp.Length - 1;

                            if (cntDelim > maxCnt)
                            {
                                maxCnt = cntDelim;
                                changeCnt++;
                            }
                        }
                    }

                    // Insert the max amount of rows needed
                    if (maxCnt > 0)
                    {
                        Excel.Range insertRange = activeSheet.Rows[row + 1 + ":" + (row + maxCnt)];
                        insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);

                        // Column iteration: Second pass
                        for (int col = startCol; col <= endCol; col++)
                        {
                            var cell = activeSheet.Cells[row, col] as Excel.Range;
                            string cellValue = null;

                            if (cell != null && cell.Value2 != null)
                            {
                                // Work with the string value
                                cellValue = cell.Value2.ToString();

                                if (!string.IsNullOrEmpty(cellValue))
                                {
                                    string[] temp = cellValue.Split(new[] { delimiter }, StringSplitOptions.None);

                                    // Single value column fill to max splits
                                    if (temp.Length == 1)
                                    {
                                        for (int fill = 0; fill <= maxCnt; fill++)
                                        {
                                            // temp[0] for single value
                                            activeSheet.Cells[row + fill, col].Value2 = temp[0].Trim();
                                        }
                                    }
                                    // Delimited value column fill with delimited values
                                    else
                                    {
                                        for (int fill = 0; fill < temp.Length; fill++)
                                        {
                                            activeSheet.Cells[row + fill, col].Value2 = temp[fill].Trim();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // Changes were made
                if (changeCnt > 0)
                {
                    // Check original range
                    if (originalRange == null)
                    {
                        originalRange = usedRange;
                    }
                    // Unwrap text
                    usedRange.WrapText = false;

                    // Select Range A1
                    Excel.Range firstCell = activeSheet.Cells[originalRange.Row, originalRange.Column];
                    firstCell.Select();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"An error occurred: {ex.Message}\nStack trace: {ex.StackTrace}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                if (usedRange != null && Marshal.IsComObject(usedRange))
                {
                    Marshal.ReleaseComObject(usedRange);
                }
                if (activeSheet != null && Marshal.IsComObject(activeSheet))
                {
                    Marshal.ReleaseComObject(activeSheet);
                }

                excelApp.ScreenUpdating = true;
            }
        }

        // OK button
        private void SplitterForm_Ok_Click(object sender, EventArgs e)
        {
            // Get the input from the TextBox
            string delimiter = CbDelimiter.Text;
            string customValue = TbCustom.Text;

            // Split to rows
            SplitToRows(_activeSheet, delimiter, customValue);
            this.Close();        
        }

        // Cancel button
        private void SplitterForm_Cancel_Click(object sender, EventArgs e)
        {
            // Select cell A1 first
            _activeSheet.Range["A1"].Select();
            this.Close();
        }

        // Delimiter Combobox change
        private void cbDelimiter_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.CbDelimiter.Text != "--Custom--")
            {
                // Clear and disable
                this.TbCustom.Text = "";
                this.TbCustom.Enabled = false;
            }
            else
            {
                // Enable
                this.TbCustom.Enabled = true;
            }
        }

        // My data has headers checkbox
        private void CbHeaders_CheckedChanged(object sender, EventArgs e)
        {
            if (usedRange == null) return;

            // Cache the original range the first time
            if (originalRange == null)
            {
                originalRange = usedRange;
            }

            if (CbHeaders.Checked)
            {
                // Only adjust if more than 1 row
                if (originalRange.Rows.Count > 1)
                {
                    usedRange = originalRange.Offset[1, 0].Resize[originalRange.Rows.Count - 1, originalRange.Columns.Count];
                }
                else
                {
                    usedRange = originalRange;
                }
            }
            else
            {
                usedRange = originalRange;
            }

            usedRange.Select();
        }

    }
}
