using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    public partial class UniqueSheetsForm : Form
    {
        private readonly Excel.Worksheet activeSheet;
        private readonly Excel.Range rangeToProcess;
        private readonly int uniqueCount;

        public UniqueSheetsForm(Excel.Range rangeToProcess, int uniqueCount)
        {
            InitializeComponent();
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            this.activeSheet = excelApp.ActiveSheet;
            this.rangeToProcess = rangeToProcess;
            this.uniqueCount = uniqueCount;
        }

        // On Load
        private void UniqueSheetsForm_Load(object sender, EventArgs e)
        {
            // Process pending Windows messages to clear the spinning cursor
            Application.DoEvents();

            // Show what will be processed
            TbUniqueValues.Text = uniqueCount.ToString("N0");

            // Nothing to create - let the user out but do not offer to run it
            UniqueSheetsForm_Ok.Enabled = uniqueCount > 0;
        }

        // OK button
        private void UniqueSheetsForm_Ok_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // Cancel button
        private void UniqueSheetsForm_Cancel_Click(object sender, EventArgs e)
        {
            // Select cell A1 to clear selection
            activeSheet.Range["A1"].Select();

            this.DialogResult = DialogResult.Cancel;
            // Close
            this.Close();
        }


    }
}