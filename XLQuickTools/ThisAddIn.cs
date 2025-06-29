using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using static XLQuickTools.QTConstants;

namespace XLQuickTools
{
    public partial class ThisAddIn
    {
        // Dictionary to track task panes for each workbook
        public Dictionary<Excel.Workbook, Tuple<CustomTaskPane, CustomTaskPane>> workbookTaskPanes =
            new Dictionary<Excel.Workbook, Tuple<CustomTaskPane, CustomTaskPane>>();

        // Dictionary to track visibility states for each workbook's task panes
        public Dictionary<Excel.Workbook, Tuple<bool, bool>> workbookPaneVisibilityStates =
            new Dictionary<Excel.Workbook, Tuple<bool, bool>>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookActivate += Application_WorkbookActivate;
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            UpdateTaskPaneVisibility(Wb);
        }

        // Toggle Missing task pane
        public void ToggleMissingTaskPaneVisibility()
        {
            ToggleTaskPaneVisibility(isMissingPane: true);
        }

        // Toggle Compare task pane
        public void ToggleCompareTaskPaneVisibility()
        {
            ToggleTaskPaneVisibility(isMissingPane: false);
        }

        // Task pane visibility
        private void ToggleTaskPaneVisibility(bool isMissingPane)
        {
            var activeWorkbook = this.Application.ActiveWorkbook;

            if (activeWorkbook != null)
            {
                if (!workbookTaskPanes.ContainsKey(activeWorkbook))
                {
                    CreateTaskPanesForWorkbook(activeWorkbook);
                }

                var panes = workbookTaskPanes[activeWorkbook];
                var visibilityStates = workbookPaneVisibilityStates[activeWorkbook];

                if (isMissingPane)
                {
                    panes.Item1.Visible = !panes.Item1.Visible; // Toggle Missing Task Pane
                    workbookPaneVisibilityStates[activeWorkbook] = Tuple.Create(panes.Item1.Visible, visibilityStates.Item2);
                }
                else
                {
                    panes.Item2.Visible = !panes.Item2.Visible; // Toggle Compare Task Pane
                    workbookPaneVisibilityStates[activeWorkbook] = Tuple.Create(visibilityStates.Item1, panes.Item2.Visible);
                }
            }
        }

        // Create missing and compare task panes
        private void CreateTaskPanesForWorkbook(Excel.Workbook workbook)
        {
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int paneWidth = Math.Max((int)(screenWidth * 0.25), MAX_PANE_WIDTH);

            // Create Missing Task Pane
            var missingUserControl = new UserControl_Missing(this.Application);
            var missingPane = this.CustomTaskPanes.Add(missingUserControl, "Find Missing Data", workbook.Windows[1]);
            missingPane.Width = paneWidth;
            missingPane.Visible = false;

            // Create Compare Task Pane
            var compareUserControl = new UserControl_Compare(this.Application);
            var comparePane = this.CustomTaskPanes.Add(compareUserControl, "Compare Worksheets", workbook.Windows[1]);
            comparePane.Width = paneWidth;
            comparePane.Visible = false;

            // Add panes to dictionaries
            workbookTaskPanes.Add(workbook, Tuple.Create(missingPane, comparePane));
            workbookPaneVisibilityStates.Add(workbook, Tuple.Create(false, false)); // Initial visibility states

            // Clean up task panes when the workbook is closed
            workbook.BeforeClose += delegate
            {
                if (workbookTaskPanes.ContainsKey(workbook))
                {
                    var panes = workbookTaskPanes[workbook];
                    this.CustomTaskPanes.Remove(panes.Item1);
                    this.CustomTaskPanes.Remove(panes.Item2);
                    workbookTaskPanes.Remove(workbook);
                    workbookPaneVisibilityStates.Remove(workbook);
                }
            };
        }

        // Update the visibility of each task pane
        private void UpdateTaskPaneVisibility(Excel.Workbook activeWorkbook)
        {
            foreach (var entry in workbookTaskPanes)
            {
                var workbook = entry.Key;
                var panes = entry.Value;
                var visibilityStates = workbookPaneVisibilityStates[workbook];

                if (workbook == activeWorkbook)
                {
                    // Restore visibility states for the active workbook
                    panes.Item1.Visible = visibilityStates.Item1;
                    panes.Item2.Visible = visibilityStates.Item2;
                }
                else
                {
                    // Save current visibility states and hide panes for inactive workbooks
                    workbookPaneVisibilityStates[workbook] = Tuple.Create(panes.Item1.Visible, panes.Item2.Visible);
                    panes.Item1.Visible = false;
                    panes.Item2.Visible = false;
                }
            }
        }

        // Shutdown
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) 
        {
            QTClipboard.Instance.Dispose();
        }

        // ------
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
