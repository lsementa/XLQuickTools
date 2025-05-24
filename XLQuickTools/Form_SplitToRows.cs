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

        public SplitterForm(Excel.Worksheet activeSheet)
        {
            InitializeComponent();
            _activeSheet = activeSheet;
        }

        // On Load
        private void SplitterForm_Load(object sender, EventArgs e)
        {

            // Show the user what the used range is
            Excel.Range usedRange = _activeSheet.UsedRange;
            usedRange.Select();

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

        // Main method to split delimted column cells to new rows
        public void SplitToRows(Excel.Worksheet activeSheet, string delimText, string customValue)
        {
            // Excel application reference
            var excelApp = Globals.ThisAddIn.Application;
            Excel.Range usedRange = activeSheet.UsedRange;
            if (usedRange == null) return;

            try
            {
                // Turn screen updating off
                excelApp.ScreenUpdating = false;

                // Get last row and last column
                int lastRow = usedRange.Rows.Count;
                int lastCol = usedRange.Columns.Count;
                // Change count
                int changeCnt = 0;

                // Get the delimiter
                string delimiter = QTUtils.GetDelimiter(delimText, customValue);

                for (int iLoop = lastRow; iLoop >= 2; iLoop--)
                {
                    int maxCnt = 0;

                    for (int i = 1; i <= lastCol; i++)
                    {
                        var cell = activeSheet.Cells[iLoop, i] as Excel.Range;
                        string cellValue = null;

                        // Check if the cell is not null and convert its value to a string
                        if (cell != null && cell.Value2 != null)
                        {
                            if (cell.Value2 is double || cell.Value2 is int)
                            {
                                // Convert numeric values to string
                                cellValue = cell.Value2.ToString();
                            }
                            else if (cell.Value2 is string)
                            {
                                // Directly use string values
                                cellValue = (string)cell.Value2;
                            }
                        }

                        // If cellValue is not null or empty
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            // Split the cell value by the delimiter
                            string[] temp = cellValue.Split(new[] { delimiter }, StringSplitOptions.None);

                            // Calculate the count of delimiters
                            int cntDelim = temp.Length - 1;

                            // Check if the current delimiter count exceeds the max count
                            if (cntDelim > maxCnt)
                            {
                                // Set the max
                                maxCnt = cntDelim;
                                changeCnt++;
                            }
                        }
                    }

                    // Insert number of rows based on max count
                    if (maxCnt > 0)
                    {
                        Excel.Range insertRange = activeSheet.Rows[iLoop + 1 + ":" + (iLoop + maxCnt)];
                        insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);

                    }

                    // Split and add values
                    for (int i = 1; i <= lastCol; i++)
                    {
                        var cell = activeSheet.Cells[iLoop, i] as Excel.Range;
                        string cellValue = null;

                        // Check if the cell is not null and its Value2 property has a value
                        if (cell != null && cell.Value2 != null)
                        {
                            // Handle different possible types (double, string, etc.)
                            if (cell.Value2 is double || cell.Value2 is int)
                            {
                                // Convert numeric values to string
                                cellValue = cell.Value2.ToString();
                            }
                            else if (cell.Value2 is string)
                            {
                                // If already a string, just assign it
                                cellValue = (string)cell.Value2;
                            }
                            else
                            {
                                // Handle other types if necessary (e.g., DateTime)
                                cellValue = cell.Value2.ToString();
                            }

                            // If the cellValue is not null or empty
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                // Split the value based on the delimiter
                                string[] temp = cellValue.Split(new[] { delimiter }, StringSplitOptions.None);

                                // Iterate through the split values
                                for (int jLoop = 0; jLoop < temp.Length; jLoop++)
                                {
                                    // Clean and trim the value
                                    string cleanedValue = temp[jLoop].Trim();

                                    // Clean
                                    cleanedValue = Clean(cleanedValue);

                                    // Assign the cleaned value back to the corresponding cell
                                    activeSheet.Cells[iLoop + jLoop, i].Value2 = cleanedValue;
                                }
                            }
                        }
                    }

                    // Fill blanks
                    if (maxCnt > 0)
                    {
                        Excel.Range fillRange = activeSheet.Rows[iLoop + 1 + ":" + (iLoop + maxCnt)];
                        QTFunctions.FillBlanks(fillRange);
                    }

                }

                // Clean up
                if (changeCnt > 0)
                {
                    // Turn text wrapping off
                    usedRange.WrapText = false;
                    // Select the first row
                    activeSheet.Cells[1, 1].Select();
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
                // Clean up COM objects
                if (usedRange != null && Marshal.IsComObject(usedRange))
                {
                    Marshal.ReleaseComObject(usedRange);
                }
                if (activeSheet != null && Marshal.IsComObject(activeSheet))
                {
                    Marshal.ReleaseComObject(activeSheet);
                }

                // Turn screenupdating back on
                excelApp.ScreenUpdating = true;
            }
        }

        // Custom method to mimic Excel's CLEAN function
        private string Clean(string input)
        {
            // Removes non-printable characters
            return new string(input.Where(c => !char.IsControl(c)).ToArray());
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
            // Select cell A1 to clear selection
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
    }
}
