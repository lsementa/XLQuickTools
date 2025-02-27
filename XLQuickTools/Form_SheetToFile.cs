﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    public partial class SheetFileForm : Form
    {
        private readonly Excel.Worksheet _activeSheet;
        public SheetFileForm(Excel.Worksheet activeSheet)
        {
            InitializeComponent();
            _activeSheet = activeSheet;
        }

        // On Load
        private void SheetFileForm_Load(object sender, EventArgs e)
        {
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

            // Create a generic file name with timestamp
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            this.TbFilename.Text = "Output_" + timestamp;

            // Populate the ComboBox with delimiter options
            this.CbExtension.Items.AddRange(new object[]
            {
                "CSV",
                "TXT"
            });

            // Set the default file extension
            this.CbExtension.SelectedItem = "CSV";

            // Default delimiter
            this.TbCustom.Text = ",";

            // Set Default contains headers
            this.CbContainsHeaders.Checked = true;

            // Put the cursor in the custom textbox
            this.TbCustom.Select();
        }

        // Cancel button
        private void FileForm_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //  OK button
        private void FileForm_Ok_Click(object sender, EventArgs e)
        {
            // Get the input from the TextBox
            string delimText = this.CbDelimiter.Text;
            string customText= this.TbCustom.Text;
            string extension = this.CbExtension.Text.ToLower();
            string fileName = this.TbFilename.Text;
            string fullFileName = $"{fileName}.{extension}";
            bool headers = this.CbContainsHeaders.Checked;
            bool quoteText = this.CbQuoteText.Checked;
            bool openFolder = this.CbOpenFolder.Checked;

            // Call the method and pass the user input
            SheetToFile(_activeSheet, delimText, customText, fullFileName, headers, quoteText, openFolder);

            this.Close();
        }

        // Sheet to file
        private void SheetToFile(Excel.Worksheet activeSheet, string delimText,
            string customText, string fileName, bool headers, bool quoteText, bool openFolder)
        {
            // Use FolderBrowserDialog to select a folder
            string savePath = QTUtils.SelectSaveFolder();
            if (string.IsNullOrEmpty(savePath))
            {
                return;
            }

            // Full file path
            string fullPath = System.IO.Path.Combine(savePath, fileName);
            // Delimiter
            string delimiter = QTUtils.GetDelimiter(delimText, customText);

            try
            {
                // Create the file using a StreamWriter
                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(fullPath, false, System.Text.Encoding.ASCII))
                {
                    // Get the used range in the active worksheet
                    Excel.Range usedRange = activeSheet.UsedRange;

                    // Check if the used range has any cells
                    if (usedRange.Count > 1)
                    {
                        object[,] valueArray = usedRange.Value2;

                        int lastRow = valueArray.GetLength(0);
                        int lastCol = valueArray.GetLength(1);

                        // Iterate through each row
                        for (int rowIndex = 1; rowIndex <= lastRow; rowIndex++)
                        {
                            List<string> rowValues = new List<string>();

                            // Iterate through each column
                            for (int colIndex = 1; colIndex <= lastCol; colIndex++)
                            {
                                // Read the value from the array instead of querying Excel for each cell
                                string cellValue = valueArray[rowIndex, colIndex]?.ToString() ?? string.Empty;

                                // Clean and trim the value
                                cellValue = QTFormat.Clean(cellValue);

                                // Determine if text qualifiers should be added
                                bool addQualifiers = !(rowIndex == 1 && headers) // Skip qualifiers for the first row if headers are enabled
                                                     && !string.IsNullOrEmpty(cellValue); // Skip qualifiers for blank cell values

                                if (addQualifiers)
                                {
                                    if (quoteText && IsText(cellValue) && !(cellValue.StartsWith("\"") && cellValue.EndsWith("\"")))
                                    {
                                        // Quote the value if it's text and not already quoted
                                        cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
                                    }
                                    else if (!quoteText && cellValue.Contains(delimiter) && !(cellValue.StartsWith("\"") && cellValue.EndsWith("\"")))
                                    {
                                        // Quote the value only if it contains the delimiter and is not already quoted
                                        cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
                                    }
                                }

                                // Add the cleaned value to the list
                                rowValues.Add(cellValue);
                            }

                            // Join the row values with the delimiter and write to the file
                            writer.WriteLine(string.Join(delimiter, rowValues));
                        }

                        // Open the folder when complete
                        if (openFolder)
                        {
                            System.Diagnostics.Process.Start("explorer.exe", savePath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error while creating the file: " + ex.Message);
            }
        }

        // Check if the value is not numeric or a date
        private bool IsText(string value)
        {
            return !decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out _) &&
                   !DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
        }

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

      
    }
}
