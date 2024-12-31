using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 

namespace XLQuickTools
{
    public partial class CSVForm : Form
    {
        private readonly Excel.Range _selectedRange;
        public CSVForm(Excel.Range selectedRange)
        {
            InitializeComponent();
            _selectedRange = selectedRange;
        }

        // OK button
        private void CSVForm_Ok_Click(object sender, EventArgs e)
        {
            // Get the input from the TextBox
            string leading = TbLeading.Text;
            string trailing = TbTrailing.Text;
            int newLine = 0;

            if (RbNewLine1.Checked)
            {
                newLine = 1; // New Lines delimiter
            }
            else if (RbNewLine2.Checked)
            {
                newLine = 2; // New lines no delimiter
            }
            else if(RbNewLine3.Checked) 
            {
                newLine = 0; // Default no new lines
            }

            // Call the method and pass the user input
            QTFunctions.SelectionPlus(leading, trailing, newLine);

            // Close after running
            this.Close();
        }

        // Close
        private void CSVForm_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // On load
        private void CSVForm_Load(object sender, EventArgs e)
        {
            this.TbLeading.Text = "'";
            this.TbTrailing.Text = "'";
            this.RbNewLine3.Checked = true;
        }
    }
}
