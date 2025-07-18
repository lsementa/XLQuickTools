﻿using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using static XLQuickTools.QTUtils;
using static XLQuickTools.QTConstants;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{

    public partial class XLQuickTools
    {
        // Tracking
        private bool isTrackingEnabled = false;

        // Remove Excess (Worksheet)
        private void BtnRemoveExcess_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.RemoveExcess();
        }

        // Remove Excess (Worksheet Button)
        private void BtnRemoveExcessWS_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.RemoveExcess();
        }

        // Remove Excess (Workbook)
        private void BtnRemoveExcessWB_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.RemoveExcess(true);
        }

        // Trim & Clean a selected range (Main button)
        private void BtnTrimClean_Click_1(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(TRIMCLEAN);
        }

        // Trim & Clean a selected range (Dropdown button)
        private void BtnTrimClean_Click_2(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(TRIMCLEAN);
        }

        // Trim & Clean active worksheet
        private void BtnTrimCleanWorksheet_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.TrimClean("worksheet");
        }

        // Trim & Clean active workbook
        private void BtnTrimCleanWorkbook_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.TrimClean("workbook");
        }

        // Trim and Clean Settings
        private void BtnTrimCleanSettings_Click(object sender, RibbonControlEventArgs e)
        {
            CleanSettingsForm form1 = new CleanSettingsForm();
            form1.ShowDialog();
        }

        // Remove formatting
        private void BtnRemoveFormatting_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet;
            QTFormat.RemoveFormattingNoUpdates(activeSheet);
        }

        // Remove formatting (Menu)
        private void BtnRemoveFormattingSub_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet;
            QTFormat.RemoveFormattingNoUpdates(activeSheet);
        }

        // Convert tables to ranges and remove all formatting
        private void BtnConvertTableRemove_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet;

            // Convert tables
            QTUtils.ConvertTablesToRanges();
            // Remove formatting
            QTFormat.RemoveFormattingNoUpdates(activeSheet);

        }

        // Remove formatting (All)
        private void BtnRemoveFormattingAll_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet;
            QTFormat.RemoveFormattingNoUpdates(activeSheet, true);
        }

        // Selection to clipboard
        private void BtnCommaSelection_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.SelectionPlus();
        }

        // Selection+ to clipboard (Form)
        private void BtnDelimSelection_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range selection = excelApp.Selection;
            CSVForm form1 = new CSVForm(selection);
            form1.ShowDialog();
        }

        // Split columns to rows (Form)
        private void BtnSplitToRows_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            SplitterForm form1 = new SplitterForm(activeSheet);
            form1.ShowDialog();
        }

        // Fill down
        private void BtnFillDown_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.FillDown();
        }

        // Worksheet to delimited file (Form)
        private void BtnSheetToFile_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            SheetFileForm form1 = new SheetFileForm(activeSheet);
            form1.ShowDialog();
        }

        // Quick Format options (Form)
        private void BtnQuickSettings_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSettingsForm form1 = new FormatSettingsForm();
            form1.ShowDialog();
        }

        // Apply Quick Format
        private void BtnQuickFormat_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.QuickFormat();
        }

        // Apply Quick Format (Menu)
        private void BtnQuickFormatSub_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.QuickFormat();
        }

        // Apply Quick Format (All)
        private void BtnQuickFormatAll_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.QuickFormat(true);
        }

        // Find duplicates
        private void BtnDuplicates_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.FindDuplicates();
        }

        // Compare worksheets (User Control)
        private void BtnCompare_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleCompareTaskPaneVisibility();
        }

        // Check for missing values (User Control)
        private void BtnMissing_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleMissingTaskPaneVisibility();
        }

        // Copy to new sheets
        private void BtnCopyToSheets_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.CopyToSheets();
        }

        // Date/Text converter (Form)
        private void BtnDateText_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            TextConvertForm form1 = new TextConvertForm(excelApp);
            form1.ShowDialog();
        }

        // Undo Date/Text formatting
        private void BtnUndo_Click(object sender, RibbonControlEventArgs e)
        {
            QTClipboard.Instance.UndoFormatting();
        }

        // Filter 
        private void BtnFilter_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.ToggleFilter();
        }

        // Set selection to uppercase
        private void BtnUppercase_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(UPPERCASE);
        }

        // Set selection to lowercase
        private void BtnLowercase_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(LOWERCASE);
        }

        // Set selection to proper case
        private void BtnPropercase_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(PROPERCASE);
        }

        // Remove letters from the selection
        private void BtnRemoveLetters_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(REMOVE_LETTERS);
        }

        // Remove numbers from the selection
        private void BtnRemoveNumbers_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(REMOVE_NUMBERS);
        }

        // Remove special characters from the selection
        private void BtnRemoveSpecial_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(REMOVE_SPECIAL);
        }

        // Normalize text
        private void BtnNormalizeText_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(NORMALIZE_TEXT);
        }

        // Add leading or trailing characters
        private void BtnAddLeadingTrailing_Click(object sender, RibbonControlEventArgs e)
        {
            LeadTrailForm form1 = new LeadTrailForm();
            form1.ShowDialog();
        }

        // Remove Non-ASCII characters
        private void BtnRemoveNonASCII_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(REPLACE_NON_ASCII);
        }

        // Remove Extra Spaces
        private void BtnRemoveSpaces_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.FormatMenu(REMOVE_SPACES);
        }

        // Subscript chemical formulas
        private void BtnSubscript_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.SubscriptChemicalFormulas();
        }

        // Delete empty rows
        private void BtnDeleteRows_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.DeleteEmptyRowsOrColumns(DeleteOption.Rows);
        }

        // Delete empty columns
        private void BtnDeleteColumns_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.DeleteEmptyRowsOrColumns(DeleteOption.Columns);
        }

        // Reset column
        private void BtnResetColumn_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.ResetColumn();
        }

        // Copy unique to clipboard (Split button)
        private void SBtnUniqueClipboard_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.UniqueSelect();
        }

        // Copy unique to clipboard
        private void BtnUniqueClipboard_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.UniqueSelect();
        }

        // Turn on/off displaying of length
        private void BtnDisplayLength_Click(object sender, RibbonControlEventArgs e)
        {
            // Toggle the tracking state
            isTrackingEnabled = !isTrackingEnabled;

            // Update the button label
            BtnDisplayLength.Label = isTrackingEnabled ? "Display &Length [On]" : "Display &Length [Off]";

            if (isTrackingEnabled)
            {
                // Subscribe to the SelectionChange event
                Globals.ThisAddIn.Application.SheetSelectionChange += Application_SheetSelectionChange;
            }
            else
            {
                // Unsubscribe from the event and clear the status bar
                Globals.ThisAddIn.Application.SheetSelectionChange -= Application_SheetSelectionChange;
                Globals.ThisAddIn.Application.StatusBar = string.Empty;
            }
        }

        // Selection change for displaying length
        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            // Only update if tracking is enabled
            if (isTrackingEnabled)
            {
                // Check if only a single cell is selected
                if (Target.Count == 1)
                {
                    // Check if the active cell contains text
                    if (Target.Value2 != null)
                    {
                        // Convert the cell value to string and calculate its length
                        string cellValue = Target.Value2.ToString();
                        int cellLength = cellValue.Length;
                        // Display the length on the status bar
                        Globals.ThisAddIn.Application.StatusBar = $"Length: {cellLength}";
                    }
                    else
                    {
                        // If the cell is empty clear status
                        Globals.ThisAddIn.Application.StatusBar = "";
                    }
                }
                else
                {
                    // If multiple cells are selected, clear the status bar
                    Globals.ThisAddIn.Application.StatusBar = "";
                }
            }
        }

        // Copy formatting to all worksheets
        private void BtnCopyFormatting_Click(object sender, RibbonControlEventArgs e)
        {
            QTFormat.CopyFormatToAllSheets();
        }

        // Create a list of sheet names
        private void BtnSheetNames_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.CopyWorksheetNamesToClipboard();
        }

        // Create a list of files
        private void BtnFileList_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            FileListForm form1 = new FileListForm(excelApp);
            form1.ShowDialog();
        }

        // Remove objects
        private void BtnRemoveObjects_Click(object sender, RibbonControlEventArgs e)
        {
            RemoveObjectsForm form1 = new RemoveObjectsForm();
            form1.ShowDialog();
        }

        // Copy highlighted cells to clipboard
        private void BtnCopyHighlightedValues_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.CopyHighlightedCellsToClipboard();
        }

        // Copy highlighted rows to new worksheet
        private void BtnCopyHighlightedRows_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            HighlightForm form1 = new HighlightForm(excelApp);
            form1.ShowDialog();
        }

        // Add custom hyperlinks - large button
        private void BtnHyperlinks_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.AddHyperlinks(true);
        }

        // Add custom hyperlinks
        private void BtnAddHyperlinks_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.AddHyperlinks(true);
        }

        // Add hyperlinks based on URL in cell
        private void BtnAddHyperlinksCell_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.AddHyperlinks(false);
        }

        // Remove hyperlinks
        private void BtnRemoveHyperlinks_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.RemoveHyperlinks();
        }

        // Hyperlink settings (Form)
        private void BtnHyperlinkSettings_Click(object sender, RibbonControlEventArgs e)
        {
            HyperlinkForm form1 = new HyperlinkForm();
            form1.ShowDialog();
        }

        // Column Information
        private void BtnColumnInfo_Click(object sender, RibbonControlEventArgs e)
        {
            QTFunctions.CountValuesInColumn();
        }
    }
}



