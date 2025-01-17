using System;
using System.Windows.Forms;

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

            // 8 = Add leading and trailing
            QTFormat.FormatMenu(8,leading, trailing);

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
