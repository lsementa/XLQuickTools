namespace XLQuickTools
{
    partial class TextConvertForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.CbFormat = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TextConvertForm_Cancel = new System.Windows.Forms.Button();
            this.TextConvertForm_Ok = new System.Windows.Forms.Button();
            this.TbExample = new System.Windows.Forms.TextBox();
            this.TbExFormatted = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.CbConvertType = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.LblFormatLocale = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.LblCurrentLocale = new System.Windows.Forms.Label();
            this.CbCurrentLocale = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.CbConvertLocale = new System.Windows.Forms.ComboBox();
            this.CbCategory = new System.Windows.Forms.ComboBox();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // CbFormat
            // 
            this.CbFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbFormat.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbFormat.FormattingEnabled = true;
            this.CbFormat.ItemHeight = 29;
            this.CbFormat.Location = new System.Drawing.Point(25, 78);
            this.CbFormat.Margin = new System.Windows.Forms.Padding(6);
            this.CbFormat.Name = "CbFormat";
            this.CbFormat.Size = new System.Drawing.Size(240, 37);
            this.CbFormat.TabIndex = 7;
            this.CbFormat.SelectedIndexChanged += new System.EventHandler(this.CbFormat_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(20, 40);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "&Category:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(20, 43);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 29);
            this.label2.TabIndex = 8;
            this.label2.Text = "&Format:";
            // 
            // TextConvertForm_Cancel
            // 
            this.TextConvertForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TextConvertForm_Cancel.Location = new System.Drawing.Point(416, 597);
            this.TextConvertForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.TextConvertForm_Cancel.Name = "TextConvertForm_Cancel";
            this.TextConvertForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.TextConvertForm_Cancel.TabIndex = 17;
            this.TextConvertForm_Cancel.Text = "Cancel";
            this.TextConvertForm_Cancel.UseVisualStyleBackColor = true;
            this.TextConvertForm_Cancel.Click += new System.EventHandler(this.TextConvertForm_Cancel_Click);
            // 
            // TextConvertForm_Ok
            // 
            this.TextConvertForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TextConvertForm_Ok.Location = new System.Drawing.Point(234, 597);
            this.TextConvertForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.TextConvertForm_Ok.Name = "TextConvertForm_Ok";
            this.TextConvertForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.TextConvertForm_Ok.TabIndex = 16;
            this.TextConvertForm_Ok.Text = "Ok";
            this.TextConvertForm_Ok.UseVisualStyleBackColor = true;
            this.TextConvertForm_Ok.Click += new System.EventHandler(this.TextConvertForm_Ok_Click);
            // 
            // TbExample
            // 
            this.TbExample.BackColor = System.Drawing.SystemColors.Window;
            this.TbExample.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbExample.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbExample.Location = new System.Drawing.Point(24, 167);
            this.TbExample.Margin = new System.Windows.Forms.Padding(4);
            this.TbExample.Name = "TbExample";
            this.TbExample.Size = new System.Drawing.Size(240, 35);
            this.TbExample.TabIndex = 13;
            this.TbExample.TextChanged += new System.EventHandler(this.TbExample_TextChanged);
            // 
            // TbExFormatted
            // 
            this.TbExFormatted.BackColor = System.Drawing.Color.Gainsboro;
            this.TbExFormatted.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbExFormatted.Enabled = false;
            this.TbExFormatted.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbExFormatted.Location = new System.Drawing.Point(294, 167);
            this.TbExFormatted.Margin = new System.Windows.Forms.Padding(4);
            this.TbExFormatted.Name = "TbExFormatted";
            this.TbExFormatted.Size = new System.Drawing.Size(240, 35);
            this.TbExFormatted.TabIndex = 15;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(20, 134);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(125, 29);
            this.label3.TabIndex = 12;
            this.label3.Text = "Test &Input:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(289, 134);
            this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(102, 29);
            this.label4.TabIndex = 14;
            this.label4.Text = "Sample:";
            // 
            // CbConvertType
            // 
            this.CbConvertType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbConvertType.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbConvertType.FormattingEnabled = true;
            this.CbConvertType.Location = new System.Drawing.Point(293, 76);
            this.CbConvertType.Name = "CbConvertType";
            this.CbConvertType.Size = new System.Drawing.Size(241, 37);
            this.CbConvertType.TabIndex = 11;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(289, 44);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(74, 29);
            this.label6.TabIndex = 10;
            this.label6.Text = "&Type:";
            // 
            // LblFormatLocale
            // 
            this.LblFormatLocale.AutoSize = true;
            this.LblFormatLocale.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblFormatLocale.Location = new System.Drawing.Point(20, 207);
            this.LblFormatLocale.Name = "LblFormatLocale";
            this.LblFormatLocale.Size = new System.Drawing.Size(180, 29);
            this.LblFormatLocale.TabIndex = 5;
            this.LblFormatLocale.Text = "C&onvert Locale:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.TbExFormatted);
            this.groupBox2.Controls.Add(this.CbFormat);
            this.groupBox2.Controls.Add(this.TbExample);
            this.groupBox2.Controls.Add(this.CbConvertType);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(22, 343);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(558, 236);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Conversion";
            // 
            // LblCurrentLocale
            // 
            this.LblCurrentLocale.AutoSize = true;
            this.LblCurrentLocale.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblCurrentLocale.Location = new System.Drawing.Point(18, 125);
            this.LblCurrentLocale.Name = "LblCurrentLocale";
            this.LblCurrentLocale.Size = new System.Drawing.Size(176, 29);
            this.LblCurrentLocale.TabIndex = 3;
            this.LblCurrentLocale.Text = "C&urrent Locale:";
            // 
            // CbCurrentLocale
            // 
            this.CbCurrentLocale.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbCurrentLocale.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbCurrentLocale.FormattingEnabled = true;
            this.CbCurrentLocale.Location = new System.Drawing.Point(23, 157);
            this.CbCurrentLocale.Name = "CbCurrentLocale";
            this.CbCurrentLocale.Size = new System.Drawing.Size(454, 37);
            this.CbCurrentLocale.TabIndex = 4;
            this.CbCurrentLocale.SelectedIndexChanged += new System.EventHandler(this.CbCurrentLocale_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CbConvertLocale);
            this.groupBox1.Controls.Add(this.CbCategory);
            this.groupBox1.Controls.Add(this.CbCurrentLocale);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.LblCurrentLocale);
            this.groupBox1.Controls.Add(this.LblFormatLocale);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(22, 21);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(558, 303);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Tag = "";
            this.groupBox1.Text = "Date/Text";
            // 
            // CbConvertLocale
            // 
            this.CbConvertLocale.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbConvertLocale.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbConvertLocale.FormattingEnabled = true;
            this.CbConvertLocale.Location = new System.Drawing.Point(23, 239);
            this.CbConvertLocale.Name = "CbConvertLocale";
            this.CbConvertLocale.Size = new System.Drawing.Size(454, 37);
            this.CbConvertLocale.TabIndex = 8;
            this.CbConvertLocale.SelectedIndexChanged += new System.EventHandler(this.CbConvertLocale_SelectedIndexChanged);
            // 
            // CbCategory
            // 
            this.CbCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbCategory.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbCategory.FormattingEnabled = true;
            this.CbCategory.Location = new System.Drawing.Point(23, 72);
            this.CbCategory.Name = "CbCategory";
            this.CbCategory.Size = new System.Drawing.Size(454, 37);
            this.CbCategory.TabIndex = 7;
            this.CbCategory.SelectedIndexChanged += new System.EventHandler(this.CbCategory_SelectedIndexChanged);
            // 
            // TextConvertForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(603, 673);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.TextConvertForm_Cancel);
            this.Controls.Add(this.TextConvertForm_Ok);
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "TextConvertForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Date/Text Converter";
            this.Load += new System.EventHandler(this.TextConvertForm_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ComboBox CbFormat;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button TextConvertForm_Cancel;
        private System.Windows.Forms.Button TextConvertForm_Ok;
        private System.Windows.Forms.TextBox TbExample;
        private System.Windows.Forms.TextBox TbExFormatted;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox CbConvertType;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label LblFormatLocale;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label LblCurrentLocale;
        private System.Windows.Forms.ComboBox CbCurrentLocale;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox CbCategory;
        private System.Windows.Forms.ComboBox CbConvertLocale;
    }
}