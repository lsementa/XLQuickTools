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
            Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
            
            if (activeSheet == null)
            {
                MessageBox.Show("Please select a worksheet.", "Remove Objects",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Turn off screen updating
                app.ScreenUpdating = false;

                for (int i = activeSheet.Shapes.Count; i >= 1; i--)
                {
                    Excel.Shape shape = activeSheet.Shapes.Item(i);

                    switch (shape.Type)
                    {
                        case Microsoft.Office.Core.MsoShapeType.msoChart:
                            if (CbCharts.Checked) shape.Delete();
                            break;

                        case Microsoft.Office.Core.MsoShapeType.msoOLEControlObject:
                        case Microsoft.Office.Core.MsoShapeType.msoEmbeddedOLEObject:
                            if (CbActiveX.Checked) shape.Delete();
                            break;

                        case Microsoft.Office.Core.MsoShapeType.msoFormControl:
                            if (CbFormControls.Checked) shape.Delete();
                            break;

                        default:
                            if (CbShapes.Checked) shape.Delete();
                            break;
                    }
                }

                // Comments
                if (CbComments.Checked)
                {
                    // Legacy notes
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