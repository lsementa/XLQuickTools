using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace XLQuickTools
{
    public partial class UniqueDataForm : Form
    {
        private readonly Excel.Worksheet activeSheet;
        private Excel.Range rangeToProcess;
        private Excel.Range originalRange;

        // Suppresses recounting while the form is still building itself
        private bool isLoading = true;

        public UniqueDataForm(Excel.Range rangeToProcess)
        {
            InitializeComponent();
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            this.activeSheet = excelApp.ActiveSheet;
            this.rangeToProcess = rangeToProcess;
        }

        // On Load
        private void UniqueDataForm_Load(object sender, EventArgs e)
        {

            // Process pending Windows messages to clear the spinning cursor
            Application.DoEvents();

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

            // Delimiter fields stay off until the checkbox is ticked
            this.CbDelimiter.Enabled = CbHasDelimiter.Checked;
            this.TbCustom.Enabled = false;

            // Recount as the custom delimiter is typed
            this.TbCustom.TextChanged += TbCustom_TextChanged;

            // Loading is done - run the counts once
            isLoading = false;
            PopulateCounts();

        }

        // Returns the active delimiter, or an empty string when the checkbox is off
        private string GetActiveDelimiter()
        {
            if (!CbHasDelimiter.Checked) return string.Empty;

            string delimText = CbDelimiter.Text;
            string customValue = TbCustom.Text;

            // Get the delimiter
            string delimiter = QTUtils.GetDelimiter(delimText, customValue);

            return delimiter ?? string.Empty;
        }

        // Has headers checkbox changed
        private void Cb_Headers_CheckedChanged(object sender, EventArgs e)
        {
            AdjustRangeAndPopulate();
        }

        // Has delimiter checkbox changed
        private void CbHasDelimiter_CheckedChanged(object sender, EventArgs e)
        {
            if (CbHasDelimiter.Checked)
            {
                // Make delimiter fields active
                CbDelimiter.Enabled = true;
                TbCustom.Enabled = (CbDelimiter.Text == "--Custom--");
            }
            else
            {
                // Make delimiter fields inactive
                CbDelimiter.Enabled = false;
                TbCustom.Enabled = false;
            }

            // Put the cursor in the custom textbox
            if (TbCustom.Enabled)
            {
                this.TbCustom.Select();
            }

            // Delimiter state changed - refresh the counts
            PopulateCounts();
        }

        // Delimiter combobox changed
        private void CbDelimiter_SelectedIndexChanged(object sender, EventArgs e)
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
                this.TbCustom.Enabled = CbHasDelimiter.Checked;
                // Put the cursor in the custom textbox
                this.TbCustom.Select();

            }

            // Delimiter changed - refresh the counts
            PopulateCounts();
        }

        // Custom delimiter typed
        private void TbCustom_TextChanged(object sender, EventArgs e)
        {
            PopulateCounts();
        }

        // Adjusts range based on checkbox state and updates UI
        private void AdjustRangeAndPopulate()
        {
            if (rangeToProcess == null) return;

            // Cache the original range the first time
            if (originalRange == null)
            {
                originalRange = rangeToProcess;
            }

            if (CbHeaders.Checked)
            {
                // Only adjust if more than 1 row
                if (originalRange.Rows.Count > 1)
                {
                    rangeToProcess = originalRange.Offset[1, 0].Resize[originalRange.Rows.Count - 1, originalRange.Columns.Count];
                }
                else
                {
                    rangeToProcess = originalRange;
                }

                // Make include headers active
                CbHeadersInclude.Enabled = true;
                CbHeadersInclude.Checked = true;
            }
            else
            {
                rangeToProcess = originalRange;

                // Make include headers inacvite
                CbHeadersInclude.Checked = false;
                CbHeadersInclude.Enabled = false;
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
                        string colLetter = QTUtils.GetColumnLetter(range.Cells[1, i].Column);
                        ClbColumns.Items.Add("Column " + colLetter, true);
                    }
                }
            }
            else
            {
                // Use column letters
                for (int i = 1; i <= colCount; i++)
                {
                    string colLetter = QTUtils.GetColumnLetter(range.Cells[1, i].Column);
                    ClbColumns.Items.Add("Column " + colLetter, true);
                }
            }

            // Get the counts
            PopulateCounts();
        }

        private void PopulateCounts()
        {
            // Skip while the form is still loading - Load does the final pass
            if (isLoading) return;

            string delimiter = GetActiveDelimiter();

            // Populate Unique Values Count
            TbUniqueValues.Text = QTFunctions.GetUniqueCount(rangeToProcess, delimiter).ToString("N0");
            TbUniqueRows.Text = QTFunctions.GetUniqueRows(rangeToProcess, ClbColumns, false, delimiter).ToString("N0");

        }

        // Checkbox Listbox Change
        private void ClbColumns_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isLoading) return;

            TbUniqueRows.Text = QTFunctions
                .GetUniqueRows(rangeToProcess, ClbColumns, false, GetActiveDelimiter())
                .ToString("N0");
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
            bool includeHeaderRow = CbHeadersInclude.Checked;

            // Include the headers when copying
            if (includeHeaderRow)
            {
                rangeToProcess = rangeToProcess.Offset[-1, 0].Resize[rangeToProcess.Rows.Count + 1, rangeToProcess.Columns.Count];
            }

            // Copy unique data to clipboard
            _ = QTFunctions.GetUniqueRows(rangeToProcess, ClbColumns, true, GetActiveDelimiter(), includeHeaderRow);

            // Close
            this.Close();
        }

        // Cancel button
        private void UniqueForm_Cancel_Click(object sender, EventArgs e)
        {
            // Select cell A1 to clear selection
            activeSheet.Range["A1"].Select();
            // Close
            this.Close();
        }

    }
}