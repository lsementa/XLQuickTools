using System;
using System.Drawing;
using System.Windows.Forms;
using static XLQuickTools.QTSettings;
using static XLQuickTools.QTConstants;

namespace XLQuickTools
{
    public partial class FormatSettingsForm : Form
    {
        public FormatSettingsForm()
        {
            InitializeComponent();
            TbZoom.Leave += TbZoom_Leave;
            TbFontSize.Leave += TbFontSize_Leave;
            TbRowHeight.Leave += TbRowHeight_Leave;
            TbColWidth.Leave += TbColWidth_Leave;
        }

        // On Load
        private void FormatSettingsForm_Load(object sender, EventArgs e)
        {
            // Load settings
            UserSettings settings = QTSettings.LoadUserSettingsFromXml();

            CbBold.Checked = settings.Bold;
            CbUnderline.Checked = settings.Underline;
            CbFreeze.Checked = settings.Freeze;
            CbBorder1.Checked = settings.Border1;
            CbBorder2.Checked = settings.Border2;
            CbInteriorColor.Checked = settings.InteriorColor;
            CbAutoFit1.Checked = settings.AutoFit1;
            CbAutoFit2.Checked = settings.AutoFit2;
            CbAutoFilter.Checked = settings.AutoFilter;
            CbWrapText.Checked = settings.WrapText;
            CbSetZoom.Checked = settings.Zoom;
            TbZoom.Text = settings.ZoomValue;
            CbGridlines.Checked = settings.Gridlines;
            TbFontSize.Text = settings.FontSizeTxt;
            CbFontSize.Checked = settings.FontSize;
            CbAllignLeft.Checked = settings.AllignLeft;
            CbAllignRight.Checked = settings.AllignRight;
            CbAllignCenter.Checked = settings.AllignCenter;
            CbAllignLeft_FR.Checked = settings.AllignLeft_FR;
            CbAllignRight_FR.Checked = settings.AllignRight_FR;
            CbAllignCenter_FR.Checked = settings.AllignCenter_FR;
            CbAllignTop.Checked = settings.AllignTop;
            CbAllignMiddle.Checked = settings.AllignMiddle;
            CbAllignBottom.Checked = settings.AllignBottom;
            CbAllignTop_FR.Checked = settings.AllignTop_FR;
            CbAllignMiddle_FR.Checked = settings.AllignMiddle_FR;
            CbAllignBottom_FR.Checked = settings.AllignBottom_FR;
            CbRowHeight.Checked = settings.RowHeight;
            CbColWidth.Checked = settings.ColWidth;
            TbRowHeight.Text = settings.Height;
            TbColWidth.Text = settings.Width;
            TbColor.Text = settings.Color;

            // Parse the OLE color stored in the TextBox (TbColor.Text)
            if (int.TryParse(TbColor.Text, out int oleColor))
            {
                // Convert the OLE color to a .NET Color object
                Color color = ColorTranslator.FromOle(oleColor);

                // Set the button's background color
                BtnSelectColor.BackColor = color;
            }
            else
            {
                BtnSelectColor.BackColor = Color.Transparent;
            }
        }

        // Save button
        private void FormatSettingsForm_Save_Click(object sender, EventArgs e)
        {
            // Load current settings from XML
            UserSettings currentSettings = QTSettings.LoadUserSettingsFromXml();

            // Update only the forms settings
            currentSettings.Bold = CbBold.Checked;
            currentSettings.Underline = CbUnderline.Checked;
            currentSettings.Freeze = CbFreeze.Checked;
            currentSettings.Border1 = CbBorder1.Checked;
            currentSettings.Border2 = CbBorder2.Checked;
            currentSettings.InteriorColor = CbInteriorColor.Checked;
            currentSettings.AutoFit1 = CbAutoFit1.Checked;
            currentSettings.AutoFit2 = CbAutoFit2.Checked;
            currentSettings.AutoFilter = CbAutoFilter.Checked;
            currentSettings.WrapText = CbWrapText.Checked;
            currentSettings.Zoom = CbSetZoom.Checked;
            currentSettings.ZoomValue = TbZoom.Text;
            currentSettings.Gridlines = CbGridlines.Checked;
            currentSettings.FontSizeTxt = TbFontSize.Text;
            currentSettings.FontSize = CbFontSize.Checked;
            currentSettings.AllignLeft = CbAllignLeft.Checked;
            currentSettings.AllignRight = CbAllignRight.Checked;
            currentSettings.AllignCenter = CbAllignCenter.Checked;
            currentSettings.AllignLeft_FR = CbAllignLeft_FR.Checked;
            currentSettings.AllignRight_FR = CbAllignRight_FR.Checked;
            currentSettings.AllignCenter_FR = CbAllignCenter_FR.Checked;
            currentSettings.AllignTop = CbAllignTop.Checked;
            currentSettings.AllignMiddle = CbAllignMiddle.Checked;
            currentSettings.AllignBottom = CbAllignBottom.Checked;
            currentSettings.AllignTop_FR = CbAllignTop_FR.Checked;
            currentSettings.AllignMiddle_FR = CbAllignMiddle_FR.Checked;
            currentSettings.AllignBottom_FR = CbAllignBottom_FR.Checked;
            currentSettings.RowHeight = CbRowHeight.Checked;
            currentSettings.ColWidth = CbColWidth.Checked;
            currentSettings.Color = TbColor.Text;
            currentSettings.Height = TbRowHeight.Text;
            currentSettings.Width = TbColWidth.Text;

            // Save all settings (checkbox + other form's values)
            QTSettings.SaveUserSettingsToXml(currentSettings);

            // Close the form
            this.Close();
        }

        // Cancel button
        private void FormatSettingsForm_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Left allignment
        private void CbAllignLeft_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignLeft.Checked)
            {
                CbAllignRight.Enabled = false;
                CbAllignCenter.Enabled = false;
            }
            else
            {
                CbAllignRight.Enabled = true;
                CbAllignCenter.Enabled = true;
            }
        }

        // Right allignment
        private void CbAllignRight_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignRight.Checked)
            {
                CbAllignLeft.Enabled = false;
                CbAllignCenter.Enabled = false;
            }
            else
            {
                CbAllignLeft.Enabled = true;
                CbAllignCenter.Enabled = true;
            }
        }

        // Center allignment
        private void CbAllignCenter_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignCenter.Checked)
            {
                CbAllignLeft.Enabled = false;
                CbAllignRight.Enabled = false;
            }
            else
            {
                CbAllignLeft.Enabled = true;
                CbAllignRight.Enabled = true;
            }
        }

        // Font Size
        private void CbFontSize_CheckedChanged(object sender, EventArgs e)
        {
            if (CbFontSize.Checked)
            {
                TbFontSize.Enabled = true;
            }
            else
            {
                TbFontSize.Clear();
                TbFontSize.Enabled = false;
            }
        }

        // Zoom
        private void CbSetZoom_CheckedChanged(object sender, EventArgs e)
        {
            if (CbSetZoom.Checked)
            {
                TbZoom.Enabled = true;
            }
            else
            {
                TbZoom.Clear();
                TbZoom.Enabled = false;
            }
        }

        // Zoom Text Change
        private void TbZoom_Leave(object sender, EventArgs e)
        {

            // Try to parse the text in the textbox to an integer
            if (int.TryParse(TbZoom.Text, out int zoomValue))
            {
                // Check thresholds
                if (zoomValue > MAX_ZOOM)
                {
                    TbZoom.Text = MAX_ZOOM.ToString();
                }
                if(zoomValue <= 0)
                {
                    TbZoom.Text = MIN_ZOOM.ToString();
                }
            }
            else
            {
                // Default if input is not valid
                TbZoom.Text = DEFAULT_ZOOM.ToString();
            }
        }


        // FontSize Text change
        private void TbFontSize_Leave(object sender, EventArgs e)
        {
            // Try to parse the text
            if (int.TryParse(TbFontSize.Text, out int fontValue))

            {
                // Check thresholds
                if (fontValue > MAX_FONT_SIZE)
                {
                    TbFontSize.Text = MAX_FONT_SIZE.ToString();
                }
                if (fontValue <= 0)
                {
                    TbFontSize.Text = DEFAULT_FONT_SIZE.ToString();
                }

            }
            else
            {
                // Default for non numeric input
                TbFontSize.Text = DEFAULT_FONT_SIZE.ToString();
            }
        }

        // Column Autofit checkbox change
        private void CbAutoFit1_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAutoFit1.Checked)
            {
                TbColWidth.Text = "";
                TbColWidth.Enabled = false;
                CbColWidth.Enabled = false;
            }
            else
            {
                CbColWidth.Enabled = true;
            }
        }

        // Rows checkbox
        private void CbAutoFit2_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAutoFit2.Checked)
            {
                TbRowHeight.Text = "";
                TbRowHeight.Enabled = false;
                CbRowHeight.Enabled = false;
            }
            else
            {
                CbRowHeight.Enabled = true;
            }
        }

        // Column Width
        private void CbColWidth_CheckedChanged(object sender, EventArgs e)
        {

            if (CbColWidth.Checked)
            {
                TbColWidth.Enabled = true;
                CbAutoFit1.Checked = false;
                CbAutoFit1.Enabled = false;
            }
            else
            {
                TbColWidth.Clear();
                TbColWidth.Enabled = false;
                CbAutoFit1.Enabled = true;
            }
            
        }

        // Row Height
        private void CbRowHeight_CheckedChanged(object sender, EventArgs e)
        {
            if (CbRowHeight.Checked)
            {
                TbRowHeight.Enabled = true;
                CbAutoFit2.Checked = false;
                CbAutoFit2.Enabled = false;
            }
            else
            {
                TbRowHeight.Clear();
                TbRowHeight.Enabled = false;
                CbAutoFit2.Enabled = true;
            }
        }

        // Row Height Text change
        private void TbRowHeight_Leave(object sender, EventArgs e)
        {
            // Try to parse the text
            if (double.TryParse(TbRowHeight.Text, out double heightValue))

            {
                // Check thresholds
                if (heightValue > MAX_ROW_HEIGHT)
                {
                    TbRowHeight.Text = MAX_ROW_HEIGHT.ToString();
                }
                if(heightValue <= 0)
                {
                    TbRowHeight.Text = DEFAULT_ROW_HEIGHT.ToString();
                }

            }
            else
            {
                // Default for non numeric input
                TbRowHeight.Text = DEFAULT_ROW_HEIGHT.ToString();
            }
        }
        // Column Width Text change
        private void TbColWidth_Leave(object sender, EventArgs e)
        {
            // Try to parse the text
            if (double.TryParse(TbColWidth.Text, out double widthValue))

            {
                // Check thresholds
                if (widthValue > MAX_COLUMN_WIDTH)
                {
                    TbColWidth.Text = MAX_COLUMN_WIDTH.ToString();
                }
                if(widthValue <= 0)
                {
                    TbColWidth.Text = DEFAULT_COLUMN_WIDTH.ToString();
                }

            }
            else
            {
                // Default for non numeric input
                TbColWidth.Text = DEFAULT_COLUMN_WIDTH.ToString();
            }
        }

        // Color selector button
        private void BtnSelectColor_Click_1(object sender, EventArgs e)
        {
            // Create and show the color picker dialog
            using (ColorDialog colorDialog = new ColorDialog())
            {
                // Optional: Allow the user to define custom colors
                colorDialog.AllowFullOpen = true;
                colorDialog.FullOpen = true;

                // Show the dialog and get the result
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    // Set the selected color to the button's background
                    Color selectedColor = colorDialog.Color;
                    BtnSelectColor.BackColor = selectedColor;

                    // Convert the selected color to OLE and store it in the textbox
                    int oleColor = ColorTranslator.ToOle(selectedColor);
                    // Store OLE color in the textbox
                    TbColor.Text = oleColor.ToString();
                }
            }
        }
        // Color selector button change
        private void CbInteriorColor_CheckedChanged(object sender, EventArgs e)
        {
            if (CbInteriorColor.Checked)
            {
                BtnSelectColor.Enabled = true;
            }
            else
            {
                TbColor.Clear();
                BtnSelectColor.Enabled = false;
                BtnSelectColor.BackColor = Color.Transparent;
            }
        }

        // Allign Left (Front Row)
        private void CbAllignLeft_FR_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignLeft_FR.Checked)
            {
                CbAllignCenter_FR.Enabled = false;
                CbAllignRight_FR.Enabled = false;
            }
            else
            {
                CbAllignCenter_FR.Enabled = true;
                CbAllignRight_FR.Enabled = true;
            }
        }

        // Allign Center (Front Row)
        private void CbAllignCenter_FR_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignCenter_FR.Checked)
            {
                CbAllignLeft_FR.Enabled = false;
                CbAllignRight_FR.Enabled = false;
            }
            else
            {
                CbAllignLeft_FR.Enabled = true;
                CbAllignRight_FR.Enabled = true;
            }
        }

        // Allign Right (Front Row)
        private void CbAllignRight_FR_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignRight_FR.Checked)
            {
                CbAllignLeft_FR.Enabled = false;
                CbAllignCenter_FR.Enabled = false;
            }
            else
            {
                CbAllignLeft_FR.Enabled = true;
                CbAllignCenter_FR.Enabled = true;
            }
        }

        // Allign Top (Front Row)
        private void CbAllignTop_FR_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignTop_FR.Checked)
            {
                CbAllignMiddle_FR.Enabled = false;
                CbAllignBottom_FR.Enabled = false;
            }
            else
            {
                CbAllignMiddle_FR.Enabled = true;
                CbAllignBottom_FR.Enabled = true;
            }
        }

        // Allign Middle (Front Row)
        private void CbAllignMiddle_FR_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignMiddle_FR.Checked)
            {
                CbAllignTop_FR.Enabled = false;
                CbAllignBottom_FR.Enabled = false;
            }
            else
            {
                CbAllignTop_FR.Enabled = true;
                CbAllignBottom_FR.Enabled = true;
            }
        }

        // Allign Bottom (Front Row)
        private void CbAllignBottom_FR_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignBottom_FR.Checked)
            {
                CbAllignTop_FR.Enabled = false;
                CbAllignMiddle_FR.Enabled = false;
            }
            else
            {
                CbAllignTop_FR.Enabled = true;
                CbAllignMiddle_FR.Enabled = true;
            }
        }

        // Allign Top
        private void CbAllignTop_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignTop.Checked)
            {
                CbAllignMiddle.Enabled = false;
                CbAllignBottom.Enabled = false;
            }
            else
            {
                CbAllignMiddle.Enabled = true;
                CbAllignBottom.Enabled = true;
            }
        }

        // Allign Middle
        private void CbAllignMiddle_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignMiddle.Checked)
            {
                CbAllignTop.Enabled = false;
                CbAllignBottom.Enabled = false;
            }
            else
            {
                CbAllignTop.Enabled = true;
                CbAllignBottom.Enabled = true;
            }
        }

        // Allign Top
        private void CbAllignBottom_CheckedChanged(object sender, EventArgs e)
        {
            if (CbAllignBottom.Checked)
            {
                CbAllignTop.Enabled = false;
                CbAllignMiddle.Enabled = false;
            }
            else
            {
                CbAllignTop.Enabled = true;
                CbAllignMiddle.Enabled = true;
            }
        }
    }
}
