using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLQuickTools
{
    public partial class FileListForm : Form
    {
        private readonly Excel.Application _excelApp;

        public FileListForm(Excel.Application excelApp)
        {
            InitializeComponent();
            _excelApp = excelApp;
        }

        // On Load
        private void FileListForm_Load(object sender, EventArgs e)
        {
            // Child Folders
            CbChildFolders.Checked = true;
            // Set to active sheet
            CbActiveSheet.Checked = true;  
            // Display settings
            CbDateModified.Checked = true;
            CbFileLocation.Checked = true;
            CbFileType.Checked = true;
            CbFileSize.Checked = true;

        }

        // Use Active Worksheet
        private void CbActiveSheet_CheckedChanged(object sender, EventArgs e)
        {
            Excel.Workbook workbook = _excelApp.ActiveWorkbook;

            if (!CbActiveSheet.Checked)
            {
                TbSheetName.Enabled = true;
                TbSheetName.Text = QTUtils.GetUniqueName("File List", workbook);
                TbSheetName.Select();
            }
            else
            {
                TbSheetName.Clear();
                TbSheetName.Enabled = false;
            }
        }

        // Browse Button
        private void FileListForm_Browse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select a folder";
                folderDialog.ShowNewFolderButton = true;

                DialogResult result = folderDialog.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderDialog.SelectedPath))
                {
                    TbFolder.Text = folderDialog.SelectedPath;
                }
            }
        }

        // Ok Button
        private void FileListForm_Ok_Click(object sender, EventArgs e)
        {
            Excel.Workbook workbook = _excelApp.ActiveWorkbook;
            Excel.Worksheet activeSheet = _excelApp.ActiveSheet;
            Excel.Worksheet sheet;

            try
            {
                // Turn off screen updating
                _excelApp.ScreenUpdating = false;

                // Get and check folder path
                string folderPath = TbFolder.Text;
                if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                {
                    MessageBox.Show("Please enter a valid folder path.", "Invalid Path", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Form options
                bool includeSubDirs = CbChildFolders.Checked;
                bool useActiveSheet = CbActiveSheet.Checked;
                string sheetName = TbSheetName.Text.Trim();
                bool includeLocation = CbFileLocation.Checked;
                bool includeDateModified = CbDateModified.Checked;
                bool includeFileType = CbFileType.Checked;
                bool includeFileSize = CbFileSize.Checked;

                // If not using active sheet confirm sheet name exists
                if (!CbActiveSheet.Checked)
                {
                    if (string.IsNullOrWhiteSpace(sheetName))
                    {
                        MessageBox.Show("Please enter a Worksheet name.", "Invalid Worksheet Name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                // Collect files
                SearchOption searchOption = includeSubDirs ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                string[] files = Directory.GetFiles(folderPath, "*.*", searchOption);

                // Get or create worksheet
                if (useActiveSheet)
                {
                    sheet = activeSheet;
                    // Clear all content and formatting
                    sheet.Cells.Clear();
                }
                else
                {
                    // Create new sheet
                    sheet = QTUtils.AddUniqueNamedWorksheet(workbook, activeSheet, sheetName);
                }

                // Header row
                int col = 1;
                sheet.Cells[1, col++] = "File Name";
                if (includeLocation) sheet.Cells[1, col++] = "File Location";
                if (includeDateModified) sheet.Cells[1, col++] = "Date Modified";
                if (includeFileType) sheet.Cells[1, col++] = "File Type";
                if (includeFileSize) sheet.Cells[1, col++] = "File Size (KB)";

                // File rows
                int row = 2;
                foreach (string filePath in files)
                {
                    FileInfo fi = new FileInfo(filePath);
                    col = 1;
                    sheet.Cells[row, col++] = fi.Name;
                    if (includeLocation) sheet.Cells[row, col++] = fi.DirectoryName;
                    if (includeDateModified) sheet.Cells[row, col++] = fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm");
                    if (includeFileType) sheet.Cells[row, col++] = fi.Extension;
                    if (includeFileSize) sheet.Cells[row, col++] = (fi.Length / 1024.0).ToString("F2"); // KB
                    row++;
                }

                // Make header row bold
                Excel.Range headerRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, col - 1]];
                headerRange.Font.Bold = true;

                // Auto-fit all used columns
                Excel.Range usedRange = sheet.UsedRange;
                usedRange.Columns.AutoFit();

                // Close the form
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Turn screen updating back on
                _excelApp.ScreenUpdating = true;
            }
        }

        // Close Button
        private void FileListForm_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
