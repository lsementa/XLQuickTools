using System;
using System.Collections.Generic;
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
        public void SplitToRows(Excel.Worksheet activeSheet, string delimText, string customValue, bool hasHeaders)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            if (usedRange == null)
            {
                System.Windows.Forms.MessageBox.Show("No data found in worksheet.", "Split Columns to Rows",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            // Set usedRange back to originalRange to be used in the process
            usedRange = originalRange;

            // Turn screen updating off
            excelApp.ScreenUpdating = false;

            try
            {
                // Read all values from the used range into a 2D array
                object[,] originalValues = (object[,])usedRange.Value2;

                if (originalValues == null)
                {
                    System.Windows.Forms.MessageBox.Show("No data found in worksheet.", "Split Columns to Rows",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // Number of rows in the fetched array
                int originalRows = originalValues.GetLength(0);
                // Number of columns in the fetched array
                int originalCols = originalValues.GetLength(1);
                // Get the delimiter
                string delimiter = QTUtils.GetDelimiter(delimText, customValue);

                List<List<object>> processedRows = new List<List<object>>();
                int totalRowsAdded = 0;
                bool changeMade = false;

                int dataStartRowIndex = 1;
                if (hasHeaders)
                {
                    // Add header row directly to processedRows list
                    List<object> headerRow = new List<object>();
                    for (int c = 1; c <= originalCols; c++)
                    {
                        headerRow.Add(originalValues[1, c]);
                    }
                    processedRows.Add(headerRow);
                    dataStartRowIndex = 2;
                    if (originalRows == 1)
                    {
                        System.Windows.Forms.MessageBox.Show("No data rows to process after skipping headers.", "Split Columns to Rows",
                           System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                        return;
                    }
                }

                // Iterate through data rows
                for (int r = dataStartRowIndex; r <= originalRows; r++)
                {
                    int maxSplitsInRow = 0;

                    // First pass: determine maximum number of delimiters used in the current row
                    for (int c = 1; c <= originalCols; c++)
                    {
                        object cellValueObj = originalValues[r, c];
                        string cellValue = cellValueObj?.ToString() ?? string.Empty;

                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            string[] parts = cellValue.Split(new[] { delimiter }, StringSplitOptions.None);
                            maxSplitsInRow = Math.Max(maxSplitsInRow, parts.Length - 1);
                        }
                    }

                    if (maxSplitsInRow == 0)
                    {
                        // No splits in this row, add it as is
                        List<object> currentRowData = new List<object>();
                        for (int c = 1; c <= originalCols; c++)
                        {
                            currentRowData.Add(originalValues[r, c]);
                        }
                        processedRows.Add(currentRowData);
                    }
                    else
                    {
                        changeMade = true;
                        totalRowsAdded += maxSplitsInRow;

                        // Initialize new rows for the current original row with empty strings
                        // +1 because if maxSplitsInRow is 1, you need 2 rows (original part + 1 split part)
                        List<List<object>> newRowsForCurrentOriginalRow = new List<List<object>>();
                        for (int k = 0; k <= maxSplitsInRow; k++)
                        {
                            newRowsForCurrentOriginalRow.Add(Enumerable.Repeat((object)string.Empty, originalCols).ToList());
                        }

                        // Second pass: populate the split data into the newRowsForCurrentOriginalRow
                        for (int c = 1; c <= originalCols; c++)
                        {
                            object cellValueObj = originalValues[r, c];
                            string cellValue = cellValueObj?.ToString() ?? string.Empty;

                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                string[] parts = cellValue.Split(new[] { delimiter }, StringSplitOptions.None)
                                                         .Select(p => p.Trim())
                                                         .ToArray();

                                // Populate the first `parts.Length` rows
                                for (int k = 0; k < parts.Length; k++)
                                {
                                    newRowsForCurrentOriginalRow[k][c - 1] = parts[k];
                                }

                                // If a column has only one part but other columns in the same row
                                // have multiple parts, fill down the single value for consistency.
                                if (parts.Length == 1 && maxSplitsInRow > 0)
                                {
                                    for (int k = 1; k <= maxSplitsInRow; k++)
                                    {
                                        newRowsForCurrentOriginalRow[k][c - 1] = parts[0];
                                    }
                                }
                            }
                        }

                        processedRows.AddRange(newRowsForCurrentOriginalRow);
                    }
                }

                if (!changeMade)
                {
                    System.Windows.Forms.MessageBox.Show("No cells contained the delimiter. No changes made.", "Split Columns to Rows",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // Convert List<List<object>> to object[,] for writing back to Excel
                object[,] finalOutputArray = new object[processedRows.Count, originalCols];
                for (int r = 0; r < processedRows.Count; r++)
                {
                    for (int c = 0; c < originalCols; c++)
                    {
                        finalOutputArray[r, c] = processedRows[r][c];
                    }
                }

                // Get the range to clear
                Excel.Range clearRange = activeSheet.Range[
                    activeSheet.Cells[originalRange.Row, originalRange.Column],
                    activeSheet.Cells[originalRange.Row + originalRows - 1 + totalRowsAdded, originalRange.Column + originalCols - 1]
                ];

                // Clear contents
                clearRange.ClearContents();

                // Determine the target range for writing the new data
                Excel.Range targetRange = activeSheet.Range[
                    activeSheet.Cells[originalRange.Row, originalRange.Column],
                    activeSheet.Cells[originalRange.Row + processedRows.Count - 1, originalRange.Column + originalCols - 1]
                ];

                // Write all processed data to Excel in one go
                targetRange.Value2 = finalOutputArray;

                // Apply formatting once at the end
                targetRange.WrapText = false;
                targetRange.ShrinkToFit = false;

                // Select Range A1 or the first cell of the original range
                Excel.Range firstCell = activeSheet.Cells[originalRange.Row, originalRange.Column];
                firstCell.Select();

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
                // Turn screenupdating back on
                excelApp.ScreenUpdating = true;
                // Clean up
                QTUtils.CleanupResources(usedRange);
                QTUtils.CleanupResources(originalRange);

            }
        }

        // OK button
        private void SplitterForm_Ok_Click(object sender, EventArgs e)
        {
            // Get the input values from the form
            string delimiter = CbDelimiter.Text;
            string customValue = TbCustom.Text;
            bool headers = CbHeaders.Checked;

            // Split to rows
            SplitToRows(_activeSheet, delimiter, customValue, headers);
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
