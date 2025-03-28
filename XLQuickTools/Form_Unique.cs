using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace XLQuickTools
{
    public partial class UniqueDataForm : Form
    {
        private Excel.Range rangeToProcess;
        private bool copyToClipboard;

        public UniqueDataForm(Excel.Range rangeToProcess, bool copyToClipboard)
        {
            InitializeComponent();
            this.rangeToProcess = rangeToProcess;
            this.copyToClipboard = copyToClipboard;
        }

        // On Load
        private void UniqueDataForm_Load(object sender, EventArgs e)
        {
            // Process pending Windows messages to clear the spinning cursor
            Application.DoEvents();

            // Change title
            this.Text = copyToClipboard ? "Selection to Clipboard" : "Selection Count";

            // Check if the range has more than 2 rows and starts at row 1
            if (rangeToProcess.Rows.Count > 2 && rangeToProcess.Row == 1)
            {
                // Automatically check there are headers
                CbHeaders.Checked = true;
            }
            else
            {
                // Disable checkbox for headers if row count = 2
                if (rangeToProcess.Rows.Count == 2)
                {
                    CbHeaders.Enabled = false;
                }
                // Populate columns list
                PopulateColumnList(rangeToProcess, false);
            }
        }

        // Checkbox Changed
        private void Cb_Headers_CheckedChanged(object sender, EventArgs e)
        {
            AdjustRangeAndPopulate();
        }

        // Adjusts range based on checkbox state and updates UI
        private void AdjustRangeAndPopulate()
        {
            if (rangeToProcess == null) return;

            Excel.Application excelApp = Globals.ThisAddIn.Application;

            if (CbHeaders.Checked)
            {
                // Exclude the first row (shift down)
                if (rangeToProcess.Rows.Count > 1)
                {
                    rangeToProcess = rangeToProcess.Offset[1, 0].Resize[rangeToProcess.Rows.Count - 1, rangeToProcess.Columns.Count];
                }
            }
            else
            {
                // Include the first row
                if (rangeToProcess.Rows.Count != 1)
                {
                    rangeToProcess = rangeToProcess.Offset[-1, 0].Resize[rangeToProcess.Rows.Count + 1, rangeToProcess.Columns.Count];
                }
            }

            // Select the adjusted range
            rangeToProcess.Select();

            // Populate columns list
            PopulateColumnList(rangeToProcess, CbHeaders.Checked);

        }

        // Populates the checked list box with column headers or column letters
        private void PopulateColumnList(Excel.Range range, bool useHeaders)
        {
            ClbColumns.Items.Clear();
            int colCount = range.Columns.Count;

            if (useHeaders)
            {
                if (range.Rows.Count > 1)
                {
                    // Get the first row as headers
                    Excel.Range headerRow = range.Worksheet.Cells[range.Row - 1, range.Column].Resize[1, colCount];
                    for (int i = 1; i <= colCount; i++)
                    {
                        string header = headerRow.Cells[1, i].Value?.ToString() ?? $"Column {i}";
                        ClbColumns.Items.Add(header, true);
                    }
                }
                else
                {
                    // Use column letters
                    for (int i = 1; i <= colCount; i++)
                    {
                        string colLetter = GetExcelColumnLetter(range.Cells[1, i].Column);
                        ClbColumns.Items.Add("Column " + colLetter, true);
                    }
                }
            }
            else
            {
                // Use column letters
                for (int i = 1; i <= colCount; i++)
                {
                    string colLetter = GetExcelColumnLetter(range.Cells[1, i].Column);
                    ClbColumns.Items.Add("Column " + colLetter, true);
                }
            }

            // Get the counts
            PopulateCounts();
        }

        private void PopulateCounts()
        {
            // Populate Unique Values Count
            TbUniqueValues.Text = QTFunctions.GetUniqueCount(rangeToProcess).ToString();
            TbUniqueRows.Text = QTFunctions.GetUniqueRows(rangeToProcess, ClbColumns, false).ToString();

        }

        // Helper function to convert column index to letter (A, B, C, ... AA, AB, etc.)
        private string GetExcelColumnLetter(int colIndex)
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

        // Checkbox Listbox Change
        private void ClbColumns_SelectedIndexChanged(object sender, EventArgs e)
        {
            TbUniqueRows.Text = QTFunctions.GetUniqueRows(rangeToProcess, ClbColumns, false).ToString();
        }

        // Select all columns
        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            SetAllChecked(true);
        }

        // Unselect all columns
        private void BtnUnselectAll_Click(object sender, EventArgs e)
        {
            SetAllChecked(false);
        }

        // Helper method to select or unselect all items
        private void SetAllChecked(bool checkAll)
        {
            for (int i = 0; i < ClbColumns.Items.Count; i++)
            {
                ClbColumns.SetItemChecked(i, checkAll);
            }
        }

        // OK button
        private void UniqueForm_Ok_Click(object sender, EventArgs e)
        {
            if (copyToClipboard)
            {
                // Copy unique data to clipboard
                _ = QTFunctions.GetUniqueRows(rangeToProcess, ClbColumns, true);
            }

            this.Close();
        }

        // Cancel button
        private void UniqueForm_Cancel_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet;
            // Select cell A1 to clear selection
            activeSheet.Range["A1"].Select();
            // Close
            this.Close();
        }

    }
}
