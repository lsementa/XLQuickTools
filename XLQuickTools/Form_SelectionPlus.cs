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
            string delimText = CbDelimiter.Text;
            string customValue = TbCustom.Text;
            int newLine = 0;

            // Get the delimiter
            string delimiter = QTUtils.GetDelimiter(delimText, customValue);

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
            QTFunctions.SelectionPlus(leading, trailing, delimiter, newLine);

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
            // Leading/Trailing and new line defaults
            this.TbLeading.Text = "'";
            this.TbTrailing.Text = "'";
            this.RbNewLine3.Checked = true;

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

            // Put the cursor in the custom textbox
            this.TbCustom.Text = ",";
            this.TbCustom.Select();
        }

        // Delimiter Combobox change
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
                this.TbCustom.Enabled = true;
            }
        }

        private void RbNewLine2_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
