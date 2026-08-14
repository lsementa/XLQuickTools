namespace XLQuickTools
{
    partial class ColumnInfoForm
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
            this.ColumnInfoForm_Ok = new System.Windows.Forms.Button();
            this.CbHeaders = new System.Windows.Forms.CheckBox();
            this.TbColumnName = new System.Windows.Forms.TextBox();
            this.TbUniqueValues = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.TbNonBlank = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.TbBlankCells = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.TbRows = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.TbDuplicateValues = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3 = new System.Windows.Forms.Form();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ColumnInfoForm_Ok
            // 
            this.ColumnInfoForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ColumnInfoForm_Ok.Location = new System.Drawing.Point(360, 305);
            this.ColumnInfoForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.ColumnInfoForm_Ok.Name = "ColumnInfoForm_Ok";
            this.ColumnInfoForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.ColumnInfoForm_Ok.TabIndex = 0;
            this.ColumnInfoForm_Ok.Text = "Ok";
            this.ColumnInfoForm_Ok.UseVisualStyleBackColor = true;
            this.ColumnInfoForm_Ok.Click += new System.EventHandler(this.ColumnInfoForm_Ok_Click);
            // 
            // CbHeaders
            // 
            this.CbHeaders.AutoSize = true;
            this.CbHeaders.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbHeaders.Location = new System.Drawing.Point(25, 318);
            this.CbHeaders.Name = "CbHeaders";
            this.CbHeaders.Size = new System.Drawing.Size(210, 33);
            this.CbHeaders.TabIndex = 12;
            this.CbHeaders.Text = "&Column header";
            this.CbHeaders.UseVisualStyleBackColor = true;
            this.CbHeaders.CheckedChanged += new System.EventHandler(this.CbHeaders_CheckedChanged);
            // 
            // TbColumnName
            // 
            this.TbColumnName.BackColor = System.Drawing.SystemColors.Window;
            this.TbColumnName.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TbColumnName.Enabled = false;
            this.TbColumnName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbColumnName.ForeColor = System.Drawing.SystemColors.WindowText;
            this.TbColumnName.Location = new System.Drawing.Point(31, 31);
            this.TbColumnName.Name = "TbColumnName";
            this.TbColumnName.Size = new System.Drawing.Size(435, 28);
            this.TbColumnName.TabIndex = 1;
            // 
            // TbUniqueValues
            // 
            this.TbUniqueValues.BackColor = System.Drawing.SystemColors.Window;
            this.TbUniqueValues.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TbUniqueValues.Enabled = false;
            this.TbUniqueValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbUniqueValues.Location = new System.Drawing.Point(287, 70);
            this.TbUniqueValues.Name = "TbUniqueValues";
            this.TbUniqueValues.Size = new System.Drawing.Size(179, 28);
            this.TbUniqueValues.TabIndex = 3;
            this.TbUniqueValues.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(51, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(175, 29);
            this.label2.TabIndex = 2;
            this.label2.Text = "Unique Values:";
            // 
            // TbNonBlank
            // 
            this.TbNonBlank.BackColor = System.Drawing.SystemColors.Window;
            this.TbNonBlank.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TbNonBlank.Enabled = false;
            this.TbNonBlank.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbNonBlank.Location = new System.Drawing.Point(287, 144);
            this.TbNonBlank.Name = "TbNonBlank";
            this.TbNonBlank.Size = new System.Drawing.Size(179, 28);
            this.TbNonBlank.TabIndex = 7;
            this.TbNonBlank.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(51, 146);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(193, 29);
            this.label3.TabIndex = 6;
            this.label3.Text = "Non-Blank Cells:";
            // 
            // TbBlankCells
            // 
            this.TbBlankCells.BackColor = System.Drawing.SystemColors.Window;
            this.TbBlankCells.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TbBlankCells.Enabled = false;
            this.TbBlankCells.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbBlankCells.Location = new System.Drawing.Point(287, 180);
            this.TbBlankCells.Name = "TbBlankCells";
            this.TbBlankCells.Size = new System.Drawing.Size(179, 28);
            this.TbBlankCells.TabIndex = 9;
            this.TbBlankCells.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(51, 182);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(140, 29);
            this.label4.TabIndex = 8;
            this.label4.Text = "Blank Cells:";
            // 
            // TbRows
            // 
            this.TbRows.BackColor = System.Drawing.SystemColors.Window;
            this.TbRows.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TbRows.Enabled = false;
            this.TbRows.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbRows.Location = new System.Drawing.Point(287, 217);
            this.TbRows.Name = "TbRows";
            this.TbRows.Size = new System.Drawing.Size(179, 28);
            this.TbRows.TabIndex = 11;
            this.TbRows.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(51, 219);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(199, 29);
            this.label5.TabIndex = 10;
            this.label5.Text = "Number of Rows:";
            // 
            // TbDuplicateValues
            // 
            this.TbDuplicateValues.BackColor = System.Drawing.SystemColors.Window;
            this.TbDuplicateValues.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TbDuplicateValues.Enabled = false;
            this.TbDuplicateValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbDuplicateValues.Location = new System.Drawing.Point(287, 107);
            this.TbDuplicateValues.Name = "TbDuplicateValues";
            this.TbDuplicateValues.Size = new System.Drawing.Size(179, 28);
            this.TbDuplicateValues.TabIndex = 5;
            this.TbDuplicateValues.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(51, 109);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(199, 29);
            this.label6.TabIndex = 4;
            this.label6.Text = "Duplicate Values:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.TbColumnName);
            this.groupBox1.Controls.Add(this.TbUniqueValues);
            this.groupBox1.Controls.Add(this.TbRows);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.TbDuplicateValues);
            this.groupBox1.Controls.Add(this.TbNonBlank);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.TbBlankCells);
            this.groupBox1.Location = new System.Drawing.Point(25, 11);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(499, 278);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            // 
            // object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3
            // 
            this.object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3.BackColor = System.Drawing.SystemColors.Control;
            this.object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3.ClientSize = new System.Drawing.Size(800, 450);
            this.object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3.Location = new System.Drawing.Point(30, 30);
            this.object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3.Name = "object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3";
            this.object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3.ShowIcon = false;
            this.object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3.Visible = false;
            // 
            // ColumnInfoForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(555, 381);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.CbHeaders);
            this.Controls.Add(this.ColumnInfoForm_Ok);
            this.Name = "ColumnInfoForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Column Information";
            this.Load += new System.EventHandler(this.ColumnInfoForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ColumnInfoForm_Ok;
        private System.Windows.Forms.CheckBox CbHeaders;
        private System.Windows.Forms.TextBox TbColumnName;
        private System.Windows.Forms.TextBox TbUniqueValues;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TbNonBlank;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox TbBlankCells;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox TbRows;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox TbDuplicateValues;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Form object_8f2df5ec_6538_4630_b9bd_f10bf440d5f3;
    }
}