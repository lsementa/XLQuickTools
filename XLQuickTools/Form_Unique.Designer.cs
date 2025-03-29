namespace XLQuickTools
{
    partial class UniqueDataForm
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
            this.ClbColumns = new System.Windows.Forms.CheckedListBox();
            this.CbHeaders = new System.Windows.Forms.CheckBox();
            this.TbUniqueValues = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TbUniqueRows = new System.Windows.Forms.TextBox();
            this.BtnSelectAll = new System.Windows.Forms.Button();
            this.BtnUnselectAll = new System.Windows.Forms.Button();
            this.UniqueForm_Ok = new System.Windows.Forms.Button();
            this.UniqueForm_Cancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ClbColumns
            // 
            this.ClbColumns.CheckOnClick = true;
            this.ClbColumns.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ClbColumns.FormattingEnabled = true;
            this.ClbColumns.Location = new System.Drawing.Point(29, 230);
            this.ClbColumns.Name = "ClbColumns";
            this.ClbColumns.Size = new System.Drawing.Size(617, 228);
            this.ClbColumns.TabIndex = 8;
            this.ClbColumns.SelectedIndexChanged += new System.EventHandler(this.ClbColumns_SelectedIndexChanged);
            // 
            // CbHeaders
            // 
            this.CbHeaders.AutoSize = true;
            this.CbHeaders.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbHeaders.Location = new System.Drawing.Point(34, 182);
            this.CbHeaders.Name = "CbHeaders";
            this.CbHeaders.Size = new System.Drawing.Size(266, 33);
            this.CbHeaders.TabIndex = 5;
            this.CbHeaders.Text = "&My data has headers";
            this.CbHeaders.UseVisualStyleBackColor = true;
            this.CbHeaders.CheckedChanged += new System.EventHandler(this.Cb_Headers_CheckedChanged);
            // 
            // TbUniqueValues
            // 
            this.TbUniqueValues.BackColor = System.Drawing.Color.Gainsboro;
            this.TbUniqueValues.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbUniqueValues.Enabled = false;
            this.TbUniqueValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbUniqueValues.Location = new System.Drawing.Point(62, 74);
            this.TbUniqueValues.Name = "TbUniqueValues";
            this.TbUniqueValues.Size = new System.Drawing.Size(216, 35);
            this.TbUniqueValues.TabIndex = 3;
            this.TbUniqueValues.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(57, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 29);
            this.label1.TabIndex = 1;
            this.label1.Text = "Values:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(331, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 29);
            this.label2.TabIndex = 2;
            this.label2.Text = "Rows:";
            // 
            // TbUniqueRows
            // 
            this.TbUniqueRows.BackColor = System.Drawing.Color.Gainsboro;
            this.TbUniqueRows.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbUniqueRows.Enabled = false;
            this.TbUniqueRows.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbUniqueRows.Location = new System.Drawing.Point(336, 74);
            this.TbUniqueRows.Name = "TbUniqueRows";
            this.TbUniqueRows.Size = new System.Drawing.Size(216, 35);
            this.TbUniqueRows.TabIndex = 4;
            this.TbUniqueRows.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // BtnSelectAll
            // 
            this.BtnSelectAll.Location = new System.Drawing.Point(348, 181);
            this.BtnSelectAll.Name = "BtnSelectAll";
            this.BtnSelectAll.Size = new System.Drawing.Size(146, 43);
            this.BtnSelectAll.TabIndex = 6;
            this.BtnSelectAll.Text = "Select &All";
            this.BtnSelectAll.UseVisualStyleBackColor = true;
            this.BtnSelectAll.Click += new System.EventHandler(this.BtnSelectAll_Click);
            // 
            // BtnUnselectAll
            // 
            this.BtnUnselectAll.Location = new System.Drawing.Point(500, 181);
            this.BtnUnselectAll.Name = "BtnUnselectAll";
            this.BtnUnselectAll.Size = new System.Drawing.Size(146, 43);
            this.BtnUnselectAll.TabIndex = 7;
            this.BtnUnselectAll.Text = "&Unselect All";
            this.BtnUnselectAll.UseVisualStyleBackColor = true;
            this.BtnUnselectAll.Click += new System.EventHandler(this.BtnUnselectAll_Click);
            // 
            // UniqueForm_Ok
            // 
            this.UniqueForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UniqueForm_Ok.Location = new System.Drawing.Point(296, 469);
            this.UniqueForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.UniqueForm_Ok.Name = "UniqueForm_Ok";
            this.UniqueForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.UniqueForm_Ok.TabIndex = 9;
            this.UniqueForm_Ok.Text = "Ok";
            this.UniqueForm_Ok.UseVisualStyleBackColor = true;
            this.UniqueForm_Ok.Click += new System.EventHandler(this.UniqueForm_Ok_Click);
            // 
            // UniqueForm_Cancel
            // 
            this.UniqueForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UniqueForm_Cancel.Location = new System.Drawing.Point(482, 469);
            this.UniqueForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.UniqueForm_Cancel.Name = "UniqueForm_Cancel";
            this.UniqueForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.UniqueForm_Cancel.TabIndex = 10;
            this.UniqueForm_Cancel.Text = "Cancel";
            this.UniqueForm_Cancel.UseVisualStyleBackColor = true;
            this.UniqueForm_Cancel.Click += new System.EventHandler(this.UniqueForm_Cancel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.TbUniqueValues);
            this.groupBox1.Controls.Add(this.TbUniqueRows);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(29, 16);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(617, 151);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Unique";
            // 
            // UniqueDataForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(675, 546);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.UniqueForm_Cancel);
            this.Controls.Add(this.UniqueForm_Ok);
            this.Controls.Add(this.BtnUnselectAll);
            this.Controls.Add(this.BtnSelectAll);
            this.Controls.Add(this.CbHeaders);
            this.Controls.Add(this.ClbColumns);
            this.Name = "UniqueDataForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Selection to Clipboard";
            this.Load += new System.EventHandler(this.UniqueDataForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox ClbColumns;
        private System.Windows.Forms.CheckBox CbHeaders;
        private System.Windows.Forms.TextBox TbUniqueValues;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TbUniqueRows;
        private System.Windows.Forms.Button BtnSelectAll;
        private System.Windows.Forms.Button BtnUnselectAll;
        private System.Windows.Forms.Button UniqueForm_Ok;
        private System.Windows.Forms.Button UniqueForm_Cancel;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}