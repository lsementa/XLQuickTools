using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    public partial class ColumnInfoForm : Form
    {
        private readonly Excel.Range columnRange;
        private bool isLoading = true;

        public ColumnInfoForm(Excel.Range columnRange)
        {
            InitializeComponent();
            this.columnRange = columnRange;
        }

        // On Load
        private void ColumnInfoForm_Load(object sender, EventArgs e)
        {
            this.CbHeaders.Checked = true;
            isLoading = false;

            UpdateColumnInfo();
        }

        // Has headers checkbox checked changed
        private void CbHeaders_CheckedChanged(object sender, EventArgs e)
        {
            if (isLoading) return;

            UpdateColumnInfo();
        }

        // Recalculate and display the counts
        private void UpdateColumnInfo()
        {
            if (columnRange == null) return;

            try
            {
                this.Cursor = Cursors.WaitCursor;

                QTFunctions.ColumnStats stats = QTFunctions.GetColumnStats(columnRange, CbHeaders.Checked);

                this.TbColumnName.Text = stats.ColumnName;
                this.TbUniqueValues.Text = stats.UniqueValues.ToString("N0");
                this.TbDuplicateValues.Text = stats.DuplicateValues.ToString("N0");
                this.TbNonBlank.Text = stats.NonBlankCells.ToString("N0");
                this.TbBlankCells.Text = stats.BlankCells.ToString("N0");
                this.TbRows.Text = stats.RowCount.ToString("N0");
            }
            catch (Exception ex)
            {
                QTUtils.ShowError(ex);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        // OK button
        private void ColumnInfoForm_Ok_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}