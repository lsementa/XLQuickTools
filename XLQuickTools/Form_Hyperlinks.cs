using System;
using System.Linq;
using System.Windows.Forms;
using static XLQuickTools.QTSettings;

namespace XLQuickTools
{
    public partial class HyperlinkForm : Form
    {
        public HyperlinkForm()
        {
            InitializeComponent();
        }

        // On Load
        private void HyperlinkForm_Load(object sender, EventArgs e)
        {
            // Load settings
            UserSettings settings = QTSettings.LoadUserSettingsFromXml();

            // Populate the listbox with saved entries
            foreach (var entry in settings.HyperlinkEntries)
            {
                LbEntries.Items.Add(entry.Name);
            }

            // Find the use this hyperlink and load it
            var matchingEntry = settings.HyperlinkEntries
                .FirstOrDefault(entry => entry.Use == true);

            // If it exists set all the fields
            if (matchingEntry != null)
            {
                TbName.Text = matchingEntry.Name;
                TbURL.Text = matchingEntry.URL;
                CbIsActive.Checked = true;

                // Set the selected index in the listbox to the matching entry
                int selectedIndex = settings.HyperlinkEntries.IndexOf(matchingEntry);
                if (selectedIndex >= 0)
                {
                    LbEntries.SelectedIndex = selectedIndex;
                }
            }

        }

        // Save button
        private void HyperlinkForm_Save_Click(object sender, EventArgs e)
        {
            // Load settings
            UserSettings settings = QTSettings.LoadUserSettingsFromXml();

            // User Input
            string name = TbName.Text.Trim();
            string url = TbURL.Text.Trim();
            bool isActive = false;
            if (CbIsActive.Checked)
            {
                isActive = true;
            }

            // Check if Name and URL fields are not empty
            if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(url))
            {
                if (isActive)
                {
                    // Iterate through the list and mark all Use = false
                    foreach (var entry in settings.HyperlinkEntries)
                    {
                        entry.Use = false;
                    }
                }

                var existingEntry = settings.HyperlinkEntries.FirstOrDefault(entry => entry.Name == TbName.Text);

                if (existingEntry != null)
                {
                    // Update existing entry
                    existingEntry.URL = TbURL.Text;
                    existingEntry.Name = TbName.Text;
                    existingEntry.Use = isActive;
                }
                else
                {
                    // Add new entry
                    settings.HyperlinkEntries.Add(new HyperlinkEntry
                    {
                        Name = TbName.Text,
                        URL = TbURL.Text,
                        Use = isActive
                    });

                }

                // Save settings
                QTSettings.SaveUserSettingsToXml(settings);
            }

            this.Close();
        }

        // Cancel button
        private void HyperlinkForm_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void LbEntries_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (LbEntries.SelectedItem != null)
            {
                string selectedName = LbEntries.SelectedItem.ToString();
                // Load settings
                UserSettings settings = QTSettings.LoadUserSettingsFromXml();

                var selectedEntry = settings.HyperlinkEntries.FirstOrDefault(entry => entry.Name == selectedName);

                if (selectedEntry != null)
                {
                    // Set the form fields
                    TbName.Text = selectedEntry.Name;
                    TbURL.Text = selectedEntry.URL;
                    if (selectedEntry.Use == true)
                    {
                        CbIsActive.Checked = true;
                    }
                    else
                    {
                        CbIsActive.Checked = false;
                    }
                }
            }
        }

        // NEW Button
        private void BtnClear_Click(object sender, EventArgs e)
        {
            TbName.Clear();
            TbURL.Clear();
            CbIsActive.Checked = false;
        }

        // ADD button
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            // Load settings
            UserSettings settings = QTSettings.LoadUserSettingsFromXml();

            // User Input
            string name = TbName.Text.Trim();
            string url = TbURL.Text.Trim();

            // Check if Name and URL fields are not empty
            if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(url))
            {
                var existingEntry = settings.HyperlinkEntries.FirstOrDefault(entry => entry.Name == name);

                if (existingEntry != null)
                {
                    // Update existing entry
                    existingEntry.URL = url;
                    existingEntry.Name = name;
                }
                else
                {
                    // Add new entry
                    settings.HyperlinkEntries.Add(new HyperlinkEntry
                    {
                        Name = name,
                        URL = url
                    });
                }

                // Save settings
                QTSettings.SaveUserSettingsToXml(settings);

                // Refresh the ListBox
                LbEntries.Items.Clear();
                foreach (var entry in settings.HyperlinkEntries)
                {
                    LbEntries.Items.Add(entry.Name);
                }

            }
            else
            {
                MessageBox.Show("Name and URL cannot be empty.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // Remove button
        private void BtnRemove_Click(object sender, EventArgs e)
        {
            // Check if an item is selected in the ListBox
            if (LbEntries.SelectedItem != null)
            {
                // Get the index of the currently selected item
                int selectedIndex = LbEntries.SelectedIndex;

                // Get the selected name
                string selectedName = LbEntries.SelectedItem.ToString();

                // Load the current settings from XML
                UserSettings settings = QTSettings.LoadUserSettingsFromXml();

                // Find the entry to remove
                var entryToRemove = settings.HyperlinkEntries.FirstOrDefault(entry => entry.Name == selectedName);
                if (entryToRemove != null)
                {


                    // Remove the entry from the list
                    settings.HyperlinkEntries.Remove(entryToRemove);

                    // Save the updated settings back to the XML
                    QTSettings.SaveUserSettingsToXml(settings);

                    // Remove the selected item from the ListBox
                    LbEntries.Items.RemoveAt(selectedIndex);

                    // Clear the form fields
                    TbName.Clear();
                    TbURL.Clear();
                    CbIsActive.Checked = false;

                    // Adjust the selection: if the removed item was not the last one, select the one below
                    if (LbEntries.Items.Count > 0)
                    {
                        if (selectedIndex >= LbEntries.Items.Count)
                        {
                            // If the removed item was the last one, select the new last item
                            LbEntries.SelectedIndex = LbEntries.Items.Count - 1;
                        }
                        else
                        {
                            // Otherwise, select the item at the same index (which is now the item below)
                            LbEntries.SelectedIndex = selectedIndex;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select an item to remove.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

    }
}
