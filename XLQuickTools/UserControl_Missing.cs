using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    public partial class UserControl_Missing : UserControl
    {
        private readonly Excel.Application _excelApp;

        public UserControl_Missing(Excel.Application excelApp)
        {
            InitializeComponent();
            _excelApp = excelApp;
        }

        // On Load
        private void UserControl_Missing_Load(object sender, EventArgs e)
        {
            // Reset fields
            ResetFields();
        }

        // Function to strip the workbook name if both are the same
        string StripWorkbookNames(string rangeText, string otherRangeText)
        {
            // Regex pattern to match [workbook name]
            string pattern = @"\[(.*?)\](.*)";

            // Extract workbook name and rest of the string for both ranges
            Match matchCurrent = Regex.Match(rangeText, pattern);
            Match matchOther = Regex.Match(otherRangeText, pattern);

            if (matchCurrent.Success && matchOther.Success)
            {
                string workbookNameCurrent = matchCurrent.Groups[1].Value;
                string workbookNameOther = matchOther.Groups[1].Value;
                string remainingTextCurrent = matchCurrent.Groups[2].Value;

                // If both workbook names are the same, use the stripped version
                if (workbookNameCurrent == workbookNameOther)
                {
                    return remainingTextCurrent.Trim().Replace("$", "").Replace("'", "");
                }
            }

            // If workbook names are different return the original text with "$" replaced
            return rangeText.Replace("$", "");
        }

        // Set up the DataGridView
        private void SetupDataGridView()
        {
            // Ranges formatted
            string range1Text = TbRange1.Text;
            string range2Text = TbRange2.Text;
            string cleanedRange1 = StripWorkbookNames(range1Text, range2Text);
            string cleanedRange2 = StripWorkbookNames(range2Text, range1Text);

            // Clear existing columns
            DgMissing.Columns.Clear();

            // Create columns for the missing items in both ranges
            DataGridViewColumn missingInRange1Column = new DataGridViewTextBoxColumn();
            missingInRange1Column.HeaderText = $"Missing in {cleanedRange1}";
            missingInRange1Column.Name = "MissingInRange1";

            DataGridViewColumn missingInRange2Column = new DataGridViewTextBoxColumn();
            missingInRange2Column.HeaderText = $"Missing in {cleanedRange2}";
            missingInRange2Column.Name = "MissingInRange2";

            // Add columns to DataGridView
            DgMissing.Columns.Add(missingInRange1Column);
            DgMissing.Columns.Add(missingInRange2Column);

            // Set to not sortable and middle allignment
            foreach (DataGridViewColumn column in DgMissing.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }

        }

        private void FindMissing(object sender, EventArgs e)
        {
            Excel.Workbook workbook = _excelApp.ActiveWorkbook;
            Excel.Worksheet activeSheet = workbook.ActiveSheet;

            if (!int.TryParse(TbMaxRows.Text, out int maxRows))
            {
                // If parsing fails
                maxRows = 25;
            }

            try
            {
                // Turn off screen updating to improve performance
                _excelApp.ScreenUpdating = false;

                // Clear Columns & Rows from the datagridview
                DgMissing.Columns.Clear();
                DgMissing.Rows.Clear();

                // Convert the string input to Excel Range using Evaluate
                Excel.Range myRange1 = (Excel.Range)_excelApp.Evaluate(TbRange1.Text);
                Excel.Range myRange2 = (Excel.Range)_excelApp.Evaluate(TbRange2.Text);

                // Check if ranges are valid
                if (myRange1 == null || myRange2 == null)
                {
                    MessageBox.Show("One of the ranges is invalid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Check for null values in arrays
                if (!(myRange1.Value2 is object[,] array1) || !(myRange2.Value2 is object[,] array2))
                {
                    MessageBox.Show("One of the ranges does not contain valid data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Hash sets for fast comparison
                HashSet<object> set1 = new HashSet<object>(array1.Cast<object>().Where(x => x != null));
                HashSet<object> set2 = new HashSet<object>(array2.Cast<object>().Where(x => x != null));

                // Track missing items
                List<object> missingInRange2 = set1.Except(set2).ToList();
                List<object> missingInRange1 = set2.Except(set1).ToList();

                // Exit early if no differences in both ranges
                if (!missingInRange2.Any() && !missingInRange1.Any())
                {
                    // Show the message only once
                    MessageBox.Show("No missing data found between ranges.", "No Missing Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Populate the DataGridView with missing items
                int rowCount = Math.Max(missingInRange2.Count, missingInRange1.Count);

                // Use a worksheet instead of datagridview if high count
                if (rowCount > maxRows || rowCount > 50)
                {
                    // Create MissingReport worksheet only if differences are found
                    string baseName = "Missing";
                    string uniqueName = baseName;
                    int counter = 1;
                    while (QTUtils.WorksheetExists(workbook, uniqueName))
                    {
                        uniqueName = baseName + counter.ToString();
                        counter++;
                    }

                    Excel.Worksheet missingReportSheet = workbook.Sheets.Add(After: activeSheet);
                    missingReportSheet.Name = uniqueName;

                    // Add headers
                    missingReportSheet.Cells[1, 1].Value = $"Missing in {TbRange1.Text.Replace("$", "")}";
                    missingReportSheet.Cells[1, 2].Value = $"Missing in {TbRange2.Text.Replace("$", "")}";

                    // Output missing values
                    for (int i = 0; i < missingInRange1.Count; i++)
                    {
                        missingReportSheet.Cells[i + 2, 1].Value = missingInRange1[i];
                    }
                    for (int i = 0; i < missingInRange2.Count; i++)
                    {
                        missingReportSheet.Cells[i + 2, 2].Value = missingInRange2[i];
                    }

                    // Apply borders and formatting
                    Excel.Range headerRange = missingReportSheet.Range["A1", "B1"];
                    headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    headerRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    missingReportSheet.Columns.AutoFit();
                    // Left-align all content in the missing report
                    Excel.Range entireSheet = missingReportSheet.UsedRange;
                    entireSheet.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                }
                else
                {
                    // Set up DataGridView and use that
                    SetupDataGridView();

                    for (int i = 0; i < rowCount; i++)
                    {
                        // Add a new row for each missing item, even if one of the ranges has fewer missing items
                        DgMissing.Rows.Add(
                            i < missingInRange1.Count ? missingInRange1[i] : null, // If available, add missing in range 1
                            i < missingInRange2.Count ? missingInRange2[i] : null  // If available, add missing in range 2
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception occurred: {ex.Message}");
                MessageBox.Show($"Error comparing ranges: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Turn screen updating back on
                _excelApp.ScreenUpdating = true;
            }
        }

        private void UcMissingClose_Click(object sender, EventArgs e)
        {
            var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (activeWorkbook != null && Globals.ThisAddIn.workbookTaskPanes.ContainsKey(activeWorkbook))
            {
                // Get the task panes for the active workbook
                var panes = Globals.ThisAddIn.workbookTaskPanes[activeWorkbook];
                var visibilityStates = Globals.ThisAddIn.workbookPaneVisibilityStates[activeWorkbook];

                // Hide the Missing Task Pane and update its visibility state
                panes.Item1.Visible = false;
                Globals.ThisAddIn.workbookPaneVisibilityStates[activeWorkbook] = Tuple.Create(false, visibilityStates.Item2);
            }
        }

        // Function to clear/reset fields
        public void ResetFields()
        {
            // Reset all relevant fields
            TbRange1.Text = string.Empty;
            TbRange2.Text = string.Empty;
            TbMaxRows.Text = "25"; // Reset to default

            // Clear any DataGridView
            if (DgMissing != null)
            {
                DgMissing.Columns.Clear();
                DgMissing.Rows.Clear();
            }
        }

        // Run button
        private void BtnCheck_Click(object sender, EventArgs e)
        {
            // Ensure both ranges are selected
            if (string.IsNullOrEmpty(TbRange1.Text) || string.IsNullOrEmpty(TbRange2.Text))
            {
                return; // Exit if either range is empty
            }
            else
            {
                FindMissing(sender, e);
            }
        }

        // Range 1 textbox clicked
        private void TbRange1_Click(object sender, EventArgs e)
        {
            Excel.Range myRange = null;
            // Display InputBox for selecting a range (Type 8 allows range selection)
            object rangeInput = _excelApp.InputBox("Please select the first range:", "Range Selector", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

            if (rangeInput is bool && (bool)rangeInput == false)
            {
                // Canceled
                return;
            }

            // Try to cast the result to an Excel.Range object
            myRange = (Excel.Range)rangeInput;

            // If it's a valid range, update the TextBox
            if (myRange != null)
            {
                TbRange1.Text = myRange.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
            }
        }

        // Range 2 textbox clicked
        private void TbRange2_Click(object sender, EventArgs e)
        {
            Excel.Range myRange = null;
            // Display InputBox for selecting a range (Type 8 allows range selection)
            object rangeInput = _excelApp.InputBox("Please select the second range:", "Range Selector", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

            if (rangeInput is bool && (bool)rangeInput == false)
            {
                // Canceled
                return;
            }

            // Try to cast the result to an Excel.Range object
            myRange = (Excel.Range)rangeInput;

            // If it's a valid range, update the TextBox
            if (myRange != null)
            {
                TbRange2.Text = myRange.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
            }
        }

        // Range 2 Change event
        private void TbRange2_TextChanged(object sender, EventArgs e)
        {
            // Ensure both ranges contain data
            if (string.IsNullOrEmpty(TbRange1.Text) || string.IsNullOrEmpty(TbRange2.Text))
            {
                return;
            }
            else
            {
                FindMissing(sender, e);
            }
        }

        // Clear/refresh button
        private void BtnClear_Click(object sender, EventArgs e)
        {
            ResetFields();
        }

    }
}
