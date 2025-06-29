using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using static XLQuickTools.QTConstants;

namespace XLQuickTools
{
    public partial class UserControl_Compare : UserControl
    {
        private readonly Excel.Application _excelApp;

        public UserControl_Compare(Excel.Application excelApp)
        {
            InitializeComponent();
            _excelApp = excelApp;

            PopulateWorkbooks();

            // Event handlers for Workbook ComboBox selections
            CbWorkbooks1.SelectedIndexChanged += CbWorkbooks1_SelectedIndexChanged;
            CbWorkbooks2.SelectedIndexChanged += CbWorkbooks2_SelectedIndexChanged;

            // Event handler for datagridview
            DgCompare.CellContentClick += DgCompare_CellContentClick;

        }

        // On load
        private void UserControl_Compare_Load(object sender, EventArgs e)
        {
            ResetFields();
        }

        // Populate both cbWorkbooks1 and cbWorkbooks2 with open workbooks
        private void PopulateWorkbooks()
        {
            CbWorkbooks1.Items.Clear();
            CbWorkbooks2.Items.Clear();

            foreach (Excel.Workbook workbook in _excelApp.Workbooks)
            {
                CbWorkbooks1.Items.Add(workbook.Name);
                CbWorkbooks2.Items.Add(workbook.Name);
            }
        }

        // Event handler for cbWorkbooks1
        private void CbWorkbooks1_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateWorksheets(CbWorkbooks1, CbWorksheets1);
        }

        // Event handler for cbWorkbooks2
        private void CbWorkbooks2_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateWorksheets(CbWorkbooks2, CbWorksheets2);
        }

        // Populates the worksheets for the selected workbook in the given combo boxes
        private void PopulateWorksheets(ComboBox workbookComboBox, ComboBox worksheetComboBox)
        {
            worksheetComboBox.Items.Clear();

            // Get the selected workbook name from the passed ComboBox (cbWorkbooks1 or cbWorkbooks2)
            string selectedWorkbookName = workbookComboBox.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selectedWorkbookName)) return;

            Excel.Workbook selectedWorkbook = null;

            // Find the selected workbook
            foreach (Excel.Workbook workbook in _excelApp.Workbooks)
            {
                if (workbook.Name == selectedWorkbookName)
                {
                    selectedWorkbook = workbook;
                    break;
                }
            }

            if (selectedWorkbook != null)
            {
                // Populate the corresponding worksheet ComboBox (cbWorksheets1 or cbWorksheets2)
                foreach (Excel.Worksheet worksheet in selectedWorkbook.Worksheets)
                {
                    worksheetComboBox.Items.Add(worksheet.Name);
                }

                // Set the first worksheet as selected
                if (worksheetComboBox.Items.Count > 0)
                {
                    worksheetComboBox.SelectedIndex = 0;
                }
            }
        }

        private void SetupDataGridView()
        {
            // Clear existing columns
            DgCompare.Columns.Clear();

            // Create and add columns for the DataGridView
            DataGridViewColumn value1Column = new DataGridViewTextBoxColumn();
            value1Column.HeaderText = $"{CbWorksheets1.SelectedItem} Value";
            DataGridViewColumn ref1Column = new DataGridViewLinkColumn();
            ref1Column.HeaderText = "Reference";
            DataGridViewColumn value2Column = new DataGridViewTextBoxColumn();
            value2Column.HeaderText = $"{CbWorksheets2.SelectedItem} Value";
            DataGridViewColumn ref2Column = new DataGridViewLinkColumn();
            ref2Column.HeaderText = "Reference";

            // Add the columns to the DataGridView
            DgCompare.Columns.Add(value1Column);
            DgCompare.Columns.Add(ref1Column);
            DgCompare.Columns.Add(value2Column);
            DgCompare.Columns.Add(ref2Column);

            // Set column styles
            foreach (DataGridViewColumn column in DgCompare.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void DgCompare_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Ensure the click is valid (ignore header clicks and invalid rows/columns)
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            // Ensure the clicked column is a link column
            if (DgCompare.Columns[e.ColumnIndex] is DataGridViewLinkColumn)
            {
                // Get the link text (which contains the cell reference)
                string linkText = DgCompare[e.ColumnIndex, e.RowIndex].Value?.ToString();

                // Link text is empty or null
                if (string.IsNullOrEmpty(linkText))
                {
                    return;
                }

                Excel.Worksheet targetSheet;
                int targetRow, targetCol;

                // Determine which worksheet and cell to activate based on the clicked column
                if (e.ColumnIndex == 1)  // First column refers to the first worksheet
                {
                    targetSheet = GetWorksheetByName(GetWorkbookByName(CbWorkbooks1.SelectedItem.ToString()), CbWorksheets1.SelectedItem.ToString());
                }
                else if (e.ColumnIndex == 3)  // Third column refers to the second worksheet
                {
                    targetSheet = GetWorksheetByName(GetWorkbookByName(CbWorkbooks2.SelectedItem.ToString()), CbWorksheets2.SelectedItem.ToString());
                }
                else
                {
                    // Invalid Column Index
                    return;
                }
                // Parse and go to link
                if (TryParseA1CellReference(linkText, out targetRow, out targetCol))
                {
                    // Get the target workbook
                    Excel.Workbook targetWorkbook = targetSheet.Parent as Excel.Workbook;
                    if (targetWorkbook != null)
                    {
                        // Ensure the target workbook is activated
                        targetWorkbook.Activate();  // Activate the workbook itself

                        // Ensure the workbook window is active and not minimized
                        if (targetWorkbook.Windows.Count > 0)
                        {
                            var targetWindow = targetWorkbook.Windows[1];

                            // If the window is minimized, restore it
                            if (targetWindow.WindowState == Excel.XlWindowState.xlMinimized)
                            {
                                targetWindow.WindowState = Excel.XlWindowState.xlNormal;  // Restore the window
                            }

                            // Activate the window
                            targetWindow.Activate();
                        }
                    }

                    // Activate the target sheet within the workbook
                    targetSheet.Activate();

                    // Use Application.Goto to bring the cell into view and select it
                    Excel.Range targetCell = targetSheet.Cells[targetRow, targetCol];
                    _excelApp.Goto(targetCell, true);  // Activate the cell and bring it into view

                    // Release target COM objects after goto
                    if (targetCell != null) Marshal.ReleaseComObject(targetCell);
                    if (targetSheet != null) Marshal.ReleaseComObject(targetSheet);
                    if (targetWorkbook != null) Marshal.ReleaseComObject(targetWorkbook);
                }
                else
                {
                    // Debug.WriteLine($"Invalid cell reference: {linkText}");
                    MessageBox.Show("Invalid cell reference in link.");
                }
            }
        }

        private bool TryParseA1CellReference(string reference, out int row, out int col)
        {
            row = col = 0;
            if (string.IsNullOrWhiteSpace(reference))
                return false;

            try
            {
                // Find the position where the numbers start (row part)
                int rowIndex = 0;
                for (int i = 0; i < reference.Length; i++)
                {
                    if (char.IsDigit(reference[i]))
                    {
                        rowIndex = i;
                        break;
                    }
                }

                // Extract column letters and row numbers
                string colPart = reference.Substring(0, rowIndex); // E.g., "C"
                string rowPart = reference.Substring(rowIndex);    // E.g., "5"

                // Convert the column part (e.g., "C") to a number (e.g., 3)
                col = 0;
                foreach (char c in colPart.ToUpper())
                {
                    col = (col * 26) + (c - 'A' + 1);
                }

                // Parse the row number
                row = int.Parse(rowPart);

                return true;
            }
            catch
            {
                return false;
            }
        }

        // Method to convert column number to column letter (e.g., 3 -> C)
        private string ConvertToA1Reference(int row, int col)
        {
            string colLetter = "";

            while (col > 0)
            {
                col--; // Adjust for zero-indexing
                colLetter = (char)('A' + col % 26) + colLetter;
                col /= 26;
            }

            return $"{colLetter}{row}";
        }

        // Compare Sheets
        private void CompareSheets(object sender, EventArgs e)
        {
            Excel.Workbook workbook = _excelApp.ActiveWorkbook;
            Excel.Worksheet activeSheet = workbook.ActiveSheet;

            // Ensure all required selections are made
            if (CbWorkbooks1.SelectedItem == null || CbWorkbooks2.SelectedItem == null ||
                CbWorksheets1.SelectedItem == null || CbWorksheets2.SelectedItem == null)
            {
                return;
            }

            // Retrieve the selected workbooks and worksheets
            Excel.Workbook workbook1 = GetWorkbookByName(CbWorkbooks1.SelectedItem.ToString());
            Excel.Workbook workbook2 = GetWorkbookByName(CbWorkbooks2.SelectedItem.ToString());
            Excel.Worksheet sheet1 = GetWorksheetByName(workbook1, CbWorksheets1.SelectedItem.ToString());
            Excel.Worksheet sheet2 = GetWorksheetByName(workbook2, CbWorksheets2.SelectedItem.ToString());

            if (sheet1 == null || sheet2 == null)
            {
                MessageBox.Show("Failed to retrieve the selected worksheets.", "Worksheet Retrieval Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            bool differencesFound = false;
            _excelApp.ScreenUpdating = false;
            _excelApp.DisplayAlerts = false;

            // Max rows for DataGridView
            int maxRows;
            if (!int.TryParse(TbMaxRows.Text, out maxRows))
            {
                maxRows = DEFAULT_DATAGRID_ROWS; // If parsing fails
            }

            List<Tuple<object, string, object, string>> diffList = new List<Tuple<object, string, object, string>>(); // To track differences
            int diffCount = 0;  // Track number of differences

            try
            {
                Excel.Range range1 = sheet1.UsedRange;
                Excel.Range range2 = sheet2.UsedRange;

                // Read the entire sheet ranges into arrays
                object[,] values1 = range1.Value2 as object[,];
                object[,] values2 = range2.Value2 as object[,];

                if (values1 == null || values2 == null)
                {
                    MessageBox.Show("One or both sheets are empty.", "Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int rows = Math.Max(values1.GetLength(0), values2.GetLength(0));
                int cols = Math.Max(values1.GetLength(1), values2.GetLength(1));

                // Iterate through rows and columns to compare the sheets
                for (int row = 1; row <= rows; row++)
                {
                    for (int col = 1; col <= cols; col++)
                    {
                        object val1 = row <= values1.GetLength(0) && col <= values1.GetLength(1) ? values1[row, col] : null;
                        object val2 = row <= values2.GetLength(0) && col <= values2.GetLength(1) ? values2[row, col] : null;

                        if (!object.Equals(val1, val2))
                        {
                            differencesFound = true;
                            diffCount++;

                            // Convert R1C1 to A1 format for the cell references
                            string cellRef1 = ConvertToA1Reference(row, col);
                            string cellRef2 = ConvertToA1Reference(row, col);

                            // Add differences to a list for processing later
                            diffList.Add(new Tuple<object, string, object, string>(
                                val1?.ToString() ?? "Empty",    // Value from Sheet 1
                                cellRef1,                      // A1 Reference for Sheet 1
                                val2?.ToString() ?? "Empty",    // Value from Sheet 2
                                cellRef2                       // A1 Reference for Sheet 2
                            ));

                            // Highlight differences
                            if (CbHighlight.Checked)
                            {
                                Excel.Range cell1 = sheet1.Cells[row, col];
                                Excel.Range cell2 = sheet2.Cells[row, col];

                                cell1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                cell2.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                            }
                        }
                    }
                }

                // After comparison, decide whether to show in DataGridView or Excel worksheet
                if (diffCount <= maxRows)
                {
                    // Populate DataGridView if differences are fewer than maxRows
                    SetupDataGridView();
                    foreach (var diff in diffList)
                    {
                        DgCompare.Rows.Add(diff.Item1, diff.Item2, diff.Item3, diff.Item4);
                    }
                }
                else
                {
                    DgCompare.Columns.Clear();

                    // Create CompareReport worksheet if there are more than maxRows differences
                    Excel.Worksheet compareSheet = QTUtils.AddUniqueNamedWorksheet(workbook, activeSheet, "Compare Report");
                    compareSheet.Cells[1, 1].Value = sheet1.Name + " Cell Contains";
                    compareSheet.Cells[1, 2].Value = "Reference";
                    compareSheet.Cells[1, 3].Value = sheet2.Name + " Cell Contains";
                    compareSheet.Cells[1, 4].Value = "Reference";

                    int compareRow = 2; // For adding differences to compareSheet

                    foreach (var diff in diffList)
                    {
                        compareSheet.Cells[compareRow, 1].Value = diff.Item1;
                        compareSheet.Cells[compareRow, 2].Value = diff.Item2;
                        compareSheet.Cells[compareRow, 3].Value = diff.Item3;
                        compareSheet.Cells[compareRow, 4].Value = diff.Item4;

                        compareRow++;
                    }

                    // Finalize the compareSheet
                    Excel.Range headerRange = compareSheet.Range["A1", "D1"];
                    headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    headerRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    compareSheet.Columns.AutoFit();
                    // Left-align all content in the compareSheet
                    Excel.Range entireSheet = compareSheet.UsedRange;
                    entireSheet.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                }

                if (!differencesFound)
                {
                    DgCompare.Columns.Clear();
                    MessageBox.Show("The sheets are identical!", "Comparison Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during the comparison: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _excelApp.DisplayAlerts = true;
                _excelApp.ScreenUpdating = true;
            }
        }

        // Method to get workbook by name
        private Excel.Workbook GetWorkbookByName(string workbookName)
        {
            foreach (Excel.Workbook wb in _excelApp.Workbooks)
            {
                if (wb.Name == workbookName)
                {
                    return wb;
                }
            }
            return null;
        }

        // Method to get worksheet by name
        private Excel.Worksheet GetWorksheetByName(Excel.Workbook workbook, string worksheetName)
        {
            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == worksheetName)
                {
                    return ws;
                }
            }
            return null;
        }

        // Close button
        private void BtnClose_Click(object sender, EventArgs e)
        {
            var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (activeWorkbook != null && Globals.ThisAddIn.workbookTaskPanes.ContainsKey(activeWorkbook))
            {
                // Get the task panes for the active workbook
                var panes = Globals.ThisAddIn.workbookTaskPanes[activeWorkbook];
                var visibilityStates = Globals.ThisAddIn.workbookPaneVisibilityStates[activeWorkbook];

                // Hide the Compare Task Pane and update its visibility state
                panes.Item2.Visible = false;
                Globals.ThisAddIn.workbookPaneVisibilityStates[activeWorkbook] = Tuple.Create(false, visibilityStates.Item1);
            }
        }

        // Run button
        private void BtnRun_Click(object sender, EventArgs e)
        {
            CompareSheets(sender, e);
        }

        // Reset
        public void ResetFields()
        {
            // Reset all relevant fields
            TbMaxRows.Text = DEFAULT_DATAGRID_ROWS.ToString(); // Reset to default
            CbWorkbooks1.Text = string.Empty;
            CbWorksheets1.Text = string.Empty;
            CbWorksheets2.Text = string.Empty;
            CbWorkbooks2.Text = string.Empty;

            // Clear any DataGridView
            if (DgCompare != null)
            {
                DgCompare.Columns.Clear();
                DgCompare.Rows.Clear();
            }
            // Reload workbooks
            PopulateWorkbooks();
        }

        // Clear / Reload
        private void BtnClear_Click(object sender, EventArgs e)
        {
            ResetFields();
        }

        // Max rows textbox leave event
        private void TbMaxRows_Leave(object sender, EventArgs e)
        {
            // Try to parse the text
            if (int.TryParse(TbMaxRows.Text, out int maxRows))
            {
                // Check thresholds
                if (maxRows > MAX_ROWS)
                {
                    TbMaxRows.Text = MAX_ROWS.ToString();
                }
                if (maxRows <= 0)
                {
                    TbMaxRows.Text = MIN_ROWS.ToString();
                }

            }
            else
            {
                // Default for non numeric input
                TbMaxRows.Text = DEFAULT_DATAGRID_ROWS.ToString();
            }
        }
    }
}