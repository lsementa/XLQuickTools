using System;
using System.Windows.Forms;
using static XLQuickTools.QTConstants;

namespace XLQuickTools
{
    public partial class LeadTrailForm : Form
    {
        public LeadTrailForm()
        {
            InitializeComponent();
        }

        // OK button
        private void LeadTrailForm_Ok_Click(object sender, EventArgs e)
        {
            // Get the input from the TextBox
            string leading = TbLeading.Text;
            string trailing = TbTrailing.Text;

            // Add leading and trailing
            QTFormat.FormatMenu(ADD_LEADING_TRAILING,leading, trailing);

            // Close after running
            this.Close();
        }

        // Close
        private void LeadTrailForm_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
