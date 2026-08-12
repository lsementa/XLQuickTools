namespace XLQuickTools
{
    partial class SheetFileForm
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
            this.TbCustom = new System.Windows.Forms.TextBox();
            this.TbFilename = new System.Windows.Forms.TextBox();
            this.CbExtension = new System.Windows.Forms.ComboBox();
            this.FileForm_Cancel = new System.Windows.Forms.Button();
            this.FileForm_Ok = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.CbDelimiter = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.CbOpenFolder = new System.Windows.Forms.CheckBox();
            this.CbQuoteText = new System.Windows.Forms.CheckBox();
            this.CbContainsHeaders = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TbCustom
            // 
            this.TbCustom.BackColor = System.Drawing.SystemColors.Window;
            this.TbCustom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbCustom.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbCustom.Location = new System.Drawing.Point(32, 138);
            this.TbCustom.Margin = new System.Windows.Forms.Padding(4);
            this.TbCustom.Name = "TbCustom";
            this.TbCustom.Size = new System.Drawing.Size(234, 35);
            this.TbCustom.TabIndex = 3;
            this.TbCustom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // TbFilename
            // 
            this.TbFilename.BackColor = System.Drawing.SystemColors.Window;
            this.TbFilename.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbFilename.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbFilename.Location = new System.Drawing.Point(32, 224);
            this.TbFilename.Margin = new System.Windows.Forms.Padding(4);
            this.TbFilename.Name = "TbFilename";
            this.TbFilename.Size = new System.Drawing.Size(447, 35);
            this.TbFilename.TabIndex = 7;
            // 
            // CbExtension
            // 
            this.CbExtension.BackColor = System.Drawing.SystemColors.Window;
            this.CbExtension.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbExtension.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbExtension.FormattingEnabled = true;
            this.CbExtension.Location = new System.Drawing.Point(289, 137);
            this.CbExtension.Margin = new System.Windows.Forms.Padding(4);
            this.CbExtension.Name = "CbExtension";
            this.CbExtension.Size = new System.Drawing.Size(190, 37);
            this.CbExtension.TabIndex = 5;
            // 
            // FileForm_Cancel
            // 
            this.FileForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FileForm_Cancel.Location = new System.Drawing.Point(315, 520);
            this.FileForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.FileForm_Cancel.Name = "FileForm_Cancel";
            this.FileForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.FileForm_Cancel.TabIndex = 13;
            this.FileForm_Cancel.Text = "Cancel";
            this.FileForm_Cancel.UseVisualStyleBackColor = true;
            this.FileForm_Cancel.Click += new System.EventHandler(this.FileForm_Cancel_Click);
            // 
            // FileForm_Ok
            // 
            this.FileForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FileForm_Ok.Location = new System.Drawing.Point(132, 520);
            this.FileForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.FileForm_Ok.Name = "FileForm_Ok";
            this.FileForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.FileForm_Ok.TabIndex = 12;
            this.FileForm_Ok.Text = "Ok";
            this.FileForm_Ok.UseVisualStyleBackColor = true;
            this.FileForm_Ok.Click += new System.EventHandler(this.FileForm_Ok_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(27, 19);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "&Delimiter:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(27, 191);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(131, 29);
            this.label2.TabIndex = 6;
            this.label2.Text = "File &Name:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(284, 104);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(121, 29);
            this.label3.TabIndex = 4;
            this.label3.Text = "File &Type:";
            // 
            // CbDelimiter
            // 
            this.CbDelimiter.BackColor = System.Drawing.SystemColors.Window;
            this.CbDelimiter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbDelimiter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbDelimiter.FormattingEnabled = true;
            this.CbDelimiter.Location = new System.Drawing.Point(32, 52);
            this.CbDelimiter.Margin = new System.Windows.Forms.Padding(4);
            this.CbDelimiter.Name = "CbDelimiter";
            this.CbDelimiter.Size = new System.Drawing.Size(447, 37);
            this.CbDelimiter.TabIndex = 1;
            this.CbDelimiter.SelectedIndexChanged += new System.EventHandler(this.CbDelimiter_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(27, 104);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(101, 29);
            this.label4.TabIndex = 2;
            this.label4.Text = "&Custom:";
            // 
            // CbOpenFolder
            // 
            this.CbOpenFolder.AutoSize = true;
            this.CbOpenFolder.BackColor = System.Drawing.SystemColors.Control;
            this.CbOpenFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbOpenFolder.Location = new System.Drawing.Point(32, 142);
            this.CbOpenFolder.Name = "CbOpenFolder";
            this.CbOpenFolder.Size = new System.Drawing.Size(343, 33);
            this.CbOpenFolder.TabIndex = 11;
            this.CbOpenFolder.Text = "&Open folder when complete";
            this.CbOpenFolder.UseVisualStyleBackColor = false;
            // 
            // CbQuoteText
            // 
            this.CbQuoteText.AutoSize = true;
            this.CbQuoteText.BackColor = System.Drawing.SystemColors.Control;
            this.CbQuoteText.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbQuoteText.Location = new System.Drawing.Point(32, 94);
            this.CbQuoteText.Name = "CbQuoteText";
            this.CbQuoteText.Size = new System.Drawing.Size(380, 33);
            this.CbQuoteText.TabIndex = 10;
            this.CbQuoteText.Text = "&Automatically add text qualifiers";
            this.CbQuoteText.UseVisualStyleBackColor = false;
            // 
            // CbContainsHeaders
            // 
            this.CbContainsHeaders.AutoSize = true;
            this.CbContainsHeaders.BackColor = System.Drawing.SystemColors.Control;
            this.CbContainsHeaders.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbContainsHeaders.Location = new System.Drawing.Point(32, 45);
            this.CbContainsHeaders.Name = "CbContainsHeaders";
            this.CbContainsHeaders.Size = new System.Drawing.Size(327, 33);
            this.CbContainsHeaders.TabIndex = 9;
            this.CbContainsHeaders.Text = "&First row contains headers";
            this.CbContainsHeaders.UseVisualStyleBackColor = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CbContainsHeaders);
            this.groupBox1.Controls.Add(this.CbQuoteText);
            this.groupBox1.Controls.Add(this.CbOpenFolder);
            this.groupBox1.Location = new System.Drawing.Point(32, 279);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(447, 219);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Options";
            // 
            // SheetFileForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(513, 602);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.CbDelimiter);
            this.Controls.Add(this.FileForm_Cancel);
            this.Controls.Add(this.FileForm_Ok);
            this.Controls.Add(this.CbExtension);
            this.Controls.Add(this.TbFilename);
            this.Controls.Add(this.TbCustom);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "SheetFileForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Worksheet to File";
            this.Load += new System.EventHandler(this.SheetFileForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TbCustom;
        private System.Windows.Forms.TextBox TbFilename;
        private System.Windows.Forms.ComboBox CbExtension;
        private System.Windows.Forms.Button FileForm_Cancel;
        private System.Windows.Forms.Button FileForm_Ok;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox CbDelimiter;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox CbOpenFolder;
        private System.Windows.Forms.CheckBox CbQuoteText;
        private System.Windows.Forms.CheckBox CbContainsHeaders;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}