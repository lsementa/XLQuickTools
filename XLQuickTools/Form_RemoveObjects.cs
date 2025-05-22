using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace XLQuickTools
{
    public partial class RemoveObjectsForm : Form
    {
        public RemoveObjectsForm()
        {
            InitializeComponent();
        }

        // On Load
        private void RemoveObjectsForm_Load(object sender, EventArgs e)
        {
            // Check all on load
            CbShapes.Checked = true;
            CbCharts.Checked = true;
            CbActiveX.Checked = true;
            CbFormControls.Checked = true;
            CbComments.Checked = true;
        }

        // Cancel Button
        private void RemoveObjectsForm_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Ok Button
        private void RemoveObjectsForm_Ok_Click(object sender, EventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = app.ActiveSheet;
            try
            {
                // Turn off screen updating
                app.ScreenUpdating = false;

                // Shapes
                if (CbShapes.Checked)
                {
                    for (int i = activeSheet.Shapes.Count; i >= 1; i--)
                    {
                        activeSheet.Shapes.Item(i).Delete();
                    }
                }

                // ActiveX
                if (CbActiveX.Checked)
                {
                    for (int i = activeSheet.OLEObjects().Count; i >= 1; i--)
                    {
                        activeSheet.OLEObjects(i).Delete();
                    }
                }

                // Form Controls
                if (CbFormControls.Checked)
                {
                    for (int i = activeSheet.DrawingObjects().Count; i >= 1; i--)
                    {
                        activeSheet.DrawingObjects(i).Delete();
                    }
                }

                // Charts
                if (CbCharts.Checked)
                {
                    for (int i = activeSheet.ChartObjects().Count; i >= 1; i--)
                    {
                        activeSheet.ChartObjects(i).Delete();
                    }
                }

                // Comments - handles both legacy and modern comments
                if (CbComments.Checked)
                {
                    // Simple approach - delete all comments at once
                    activeSheet.Cells.ClearComments();
                }

                // Close form
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }
    }
}
