namespace XLQuickTools
{
    partial class CSVForm
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
            this.TbLeading = new System.Windows.Forms.TextBox();
            this.CSVForm_Ok = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.CSVForm_Cancel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.TbTrailing = new System.Windows.Forms.TextBox();
            this.RbNewLine1 = new System.Windows.Forms.RadioButton();
            this.RbNewLine2 = new System.Windows.Forms.RadioButton();
            this.RbNewLine3 = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.CbDelimiter = new System.Windows.Forms.ComboBox();
            this.TbCustom = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TbLeading
            // 
            this.TbLeading.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbLeading.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbLeading.Location = new System.Drawing.Point(33, 60);
            this.TbLeading.Margin = new System.Windows.Forms.Padding(4);
            this.TbLeading.Name = "TbLeading";
            this.TbLeading.Size = new System.Drawing.Size(384, 35);
            this.TbLeading.TabIndex = 1;
            this.TbLeading.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // CSVForm_Ok
            // 
            this.CSVForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CSVForm_Ok.Location = new System.Drawing.Point(483, 299);
            this.CSVForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.CSVForm_Ok.Name = "CSVForm_Ok";
            this.CSVForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.CSVForm_Ok.TabIndex = 12;
            this.CSVForm_Ok.Text = "Ok";
            this.CSVForm_Ok.UseVisualStyleBackColor = true;
            this.CSVForm_Ok.Click += new System.EventHandler(this.CSVForm_Ok_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(28, 27);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(297, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "&Leading Character or Text:";
            // 
            // CSVForm_Cancel
            // 
            this.CSVForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CSVForm_Cancel.Location = new System.Drawing.Point(666, 299);
            this.CSVForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.CSVForm_Cancel.Name = "CSVForm_Cancel";
            this.CSVForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.CSVForm_Cancel.TabIndex = 13;
            this.CSVForm_Cancel.Text = "Cancel";
            this.CSVForm_Cancel.UseVisualStyleBackColor = true;
            this.CSVForm_Cancel.Click += new System.EventHandler(this.CSVForm_Cancel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(29, 111);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(292, 29);
            this.label2.TabIndex = 2;
            this.label2.Text = "&Trailing Character or Text:";
            // 
            // TbTrailing
            // 
            this.TbTrailing.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbTrailing.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbTrailing.Location = new System.Drawing.Point(34, 144);
            this.TbTrailing.Margin = new System.Windows.Forms.Padding(4);
            this.TbTrailing.Name = "TbTrailing";
            this.TbTrailing.Size = new System.Drawing.Size(384, 35);
            this.TbTrailing.TabIndex = 3;
            this.TbTrailing.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // RbNewLine1
            // 
            this.RbNewLine1.AutoSize = true;
            this.RbNewLine1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RbNewLine1.Location = new System.Drawing.Point(48, 35);
            this.RbNewLine1.Margin = new System.Windows.Forms.Padding(4);
            this.RbNewLine1.Name = "RbNewLine1";
            this.RbNewLine1.Size = new System.Drawing.Size(287, 33);
            this.RbNewLine1.TabIndex = 9;
            this.RbNewLine1.TabStop = true;
            this.RbNewLine1.Text = "Include (&with delimiter)";
            this.RbNewLine1.UseVisualStyleBackColor = true;
            // 
            // RbNewLine2
            // 
            this.RbNewLine2.AutoSize = true;
            this.RbNewLine2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RbNewLine2.Location = new System.Drawing.Point(48, 70);
            this.RbNewLine2.Margin = new System.Windows.Forms.Padding(4);
            this.RbNewLine2.Name = "RbNewLine2";
            this.RbNewLine2.Size = new System.Drawing.Size(271, 33);
            this.RbNewLine2.TabIndex = 10;
            this.RbNewLine2.TabStop = true;
            this.RbNewLine2.Text = "Include (&no delimiter)";
            this.RbNewLine2.UseVisualStyleBackColor = true;
            // 
            // RbNewLine3
            // 
            this.RbNewLine3.AutoSize = true;
            this.RbNewLine3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RbNewLine3.Location = new System.Drawing.Point(48, 109);
            this.RbNewLine3.Margin = new System.Windows.Forms.Padding(4);
            this.RbNewLine3.Name = "RbNewLine3";
            this.RbNewLine3.Size = new System.Drawing.Size(198, 33);
            this.RbNewLine3.TabIndex = 11;
            this.RbNewLine3.TabStop = true;
            this.RbNewLine3.Text = "&Do not include";
            this.RbNewLine3.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.RbNewLine1);
            this.groupBox1.Controls.Add(this.RbNewLine3);
            this.groupBox1.Controls.Add(this.RbNewLine2);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(446, 111);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(384, 166);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "New Lines";
            // 
            // CbDelimiter
            // 
            this.CbDelimiter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbDelimiter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbDelimiter.FormattingEnabled = true;
            this.CbDelimiter.Location = new System.Drawing.Point(33, 227);
            this.CbDelimiter.Margin = new System.Windows.Forms.Padding(4);
            this.CbDelimiter.Name = "CbDelimiter";
            this.CbDelimiter.Size = new System.Drawing.Size(383, 37);
            this.CbDelimiter.TabIndex = 5;
            this.CbDelimiter.SelectedIndexChanged += new System.EventHandler(this.CbDelimiter_SelectedIndexChanged);
            // 
            // TbCustom
            // 
            this.TbCustom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbCustom.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbCustom.Location = new System.Drawing.Point(446, 60);
            this.TbCustom.Margin = new System.Windows.Forms.Padding(4);
            this.TbCustom.Name = "TbCustom";
            this.TbCustom.Size = new System.Drawing.Size(384, 35);
            this.TbCustom.TabIndex = 7;
            this.TbCustom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(28, 194);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(116, 29);
            this.label3.TabIndex = 4;
            this.label3.Text = "&Delimiter:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(441, 27);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(204, 29);
            this.label4.TabIndex = 6;
            this.label4.Text = "&Custom Delimiter:";
            // 
            // CSVForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(864, 382);
            this.Controls.Add(this.CbDelimiter);
            this.Controls.Add(this.TbCustom);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.TbTrailing);
            this.Controls.Add(this.CSVForm_Cancel);
            this.Controls.Add(this.CSVForm_Ok);
            this.Controls.Add(this.TbLeading);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label4);
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "CSVForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Selection+";
            this.Load += new System.EventHandler(this.CSVForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TbLeading;
        private System.Windows.Forms.Button CSVForm_Ok;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button CSVForm_Cancel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TbTrailing;
        private System.Windows.Forms.RadioButton RbNewLine1;
        private System.Windows.Forms.RadioButton RbNewLine2;
        private System.Windows.Forms.RadioButton RbNewLine3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox CbDelimiter;
        private System.Windows.Forms.TextBox TbCustom;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}