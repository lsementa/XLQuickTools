using System;
using System.Diagnostics;
using System.Windows.Forms;
using static XLQuickTools.QTSettings;

namespace XLQuickTools
{
    public partial class CleanSettingsForm : Form
    {
        private bool _loading;

        public CleanSettingsForm()
        {
            InitializeComponent();
        }

        // On load
        private void Form_TrimCleanSettings_Load(object sender, EventArgs e)
        {
            _loading = true;

            try
            {
                // Load settings
                UserSettings settings = QTSettings.LoadUserSettingsFromXml();

                CbSpaces.Checked = settings.Spaces;
                CbSpacesAll.Checked = settings.SpacesAll;
                CbNonPrintable.Checked = settings.NonPrintable;
                CbNonASCII.Checked = settings.NonASCII;

                // Apply enable/disable once, from the final loaded state
                ApplySpaceRules();
            }
            finally
            {
                _loading = false;
            }
        }

        private void ApplySpaceRules()
        {
            if (CbSpaces.Checked)
            {
                CbSpacesAll.Checked = false;
                CbSpacesAll.Enabled = false;
                CbSpaces.Enabled = true;
            }
            else if (CbSpacesAll.Checked)
            {
                CbSpaces.Enabled = false;
                CbSpacesAll.Enabled = true;
            }
            else
            {
                CbSpaces.Enabled = true;
                CbSpacesAll.Enabled = true;
            }
        }

        // Save button
        private void TrimCleanForm_Save_Click(object sender, EventArgs e)
        {
            // Load current settings from XML
            UserSettings currentSettings = QTSettings.LoadUserSettingsFromXml();

            // Update only the forms settings
            currentSettings.Spaces = CbSpaces.Checked;
            currentSettings.SpacesAll = CbSpacesAll.Checked;
            currentSettings.NonPrintable = CbNonPrintable.Checked;
            currentSettings.NonASCII = CbNonASCII.Checked;

            // Save all settings (checkbox + other form's values)
            QTSettings.SaveUserSettingsToXml(currentSettings);

            // Close the form
            this.Close();
        }

        // Cancel button
        private void TrimCleanForm_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Checkbox event for two or more spaces
        private void CbSpaces_CheckedChanged(object sender, EventArgs e)
        {
            if (_loading) return;

            if (CbSpaces.Checked)
            {
                CbSpacesAll.Checked = false;
                CbSpacesAll.Enabled = false;
            }
            else
            {
                CbSpacesAll.Enabled = true;
            }
        }

        // Checkbox event for all spaces
        private void CbSpacesAll_CheckedChanged(object sender, EventArgs e)
        {
            if (_loading) return;

            if (CbSpacesAll.Checked)
            {
                CbSpaces.Checked = false;
                CbSpaces.Enabled = false;
            }
            else
            {
                CbSpaces.Enabled = true;
            }
        }
    }
}