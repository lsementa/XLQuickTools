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
            this.label3 = new System.Windows.Forms.Label();
            this.UniqueForm_Ok = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ClbColumns
            // 
            this.ClbColumns.CheckOnClick = true;
            this.ClbColumns.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ClbColumns.FormattingEnabled = true;
            this.ClbColumns.Location = new System.Drawing.Point(29, 120);
            this.ClbColumns.Name = "ClbColumns";
            this.ClbColumns.Size = new System.Drawing.Size(617, 228);
            this.ClbColumns.TabIndex = 4;
            this.ClbColumns.SelectedIndexChanged += new System.EventHandler(this.ClbColumns_SelectedIndexChanged);
            // 
            // CbHeaders
            // 
            this.CbHeaders.AutoSize = true;
            this.CbHeaders.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbHeaders.Location = new System.Drawing.Point(29, 27);
            this.CbHeaders.Name = "CbHeaders";
            this.CbHeaders.Size = new System.Drawing.Size(266, 33);
            this.CbHeaders.TabIndex = 0;
            this.CbHeaders.Text = "&My data has headers";
            this.CbHeaders.UseVisualStyleBackColor = true;
            this.CbHeaders.CheckedChanged += new System.EventHandler(this.Cb_Headers_CheckedChanged);
            // 
            // TbUniqueValues
            // 
            this.TbUniqueValues.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbUniqueValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbUniqueValues.Location = new System.Drawing.Point(197, 365);
            this.TbUniqueValues.Name = "TbUniqueValues";
            this.TbUniqueValues.Size = new System.Drawing.Size(129, 35);
            this.TbUniqueValues.TabIndex = 6;
            this.TbUniqueValues.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(24, 367);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(171, 29);
            this.label1.TabIndex = 5;
            this.label1.Text = "Unique values:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(361, 367);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(154, 29);
            this.label2.TabIndex = 7;
            this.label2.Text = "Unique rows:";
            // 
            // TbUniqueRows
            // 
            this.TbUniqueRows.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbUniqueRows.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbUniqueRows.Location = new System.Drawing.Point(517, 365);
            this.TbUniqueRows.Name = "TbUniqueRows";
            this.TbUniqueRows.Size = new System.Drawing.Size(129, 35);
            this.TbUniqueRows.TabIndex = 8;
            this.TbUniqueRows.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // BtnSelectAll
            // 
            this.BtnSelectAll.Location = new System.Drawing.Point(348, 71);
            this.BtnSelectAll.Name = "BtnSelectAll";
            this.BtnSelectAll.Size = new System.Drawing.Size(146, 43);
            this.BtnSelectAll.TabIndex = 2;
            this.BtnSelectAll.Text = "Select &All";
            this.BtnSelectAll.UseVisualStyleBackColor = true;
            this.BtnSelectAll.Click += new System.EventHandler(this.BtnSelectAll_Click);
            // 
            // BtnUnselectAll
            // 
            this.BtnUnselectAll.Location = new System.Drawing.Point(500, 71);
            this.BtnUnselectAll.Name = "BtnUnselectAll";
            this.BtnUnselectAll.Size = new System.Drawing.Size(146, 43);
            this.BtnUnselectAll.TabIndex = 3;
            this.BtnUnselectAll.Text = "&Unselect All";
            this.BtnUnselectAll.UseVisualStyleBackColor = true;
            this.BtnUnselectAll.Click += new System.EventHandler(this.BtnUnselectAll_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(24, 85);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(114, 29);
            this.label3.TabIndex = 1;
            this.label3.Text = "Columns:";
            // 
            // UniqueForm_Ok
            // 
            this.UniqueForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UniqueForm_Ok.Location = new System.Drawing.Point(482, 430);
            this.UniqueForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.UniqueForm_Ok.Name = "UniqueForm_Ok";
            this.UniqueForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.UniqueForm_Ok.TabIndex = 9;
            this.UniqueForm_Ok.Text = "Ok";
            this.UniqueForm_Ok.UseVisualStyleBackColor = true;
            this.UniqueForm_Ok.Click += new System.EventHandler(this.UniqueForm_Ok_Click);
            // 
            // UniqueDataForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(677, 508);
            this.Controls.Add(this.UniqueForm_Ok);
            this.Controls.Add(this.BtnUnselectAll);
            this.Controls.Add(this.BtnSelectAll);
            this.Controls.Add(this.CbHeaders);
            this.Controls.Add(this.TbUniqueRows);
            this.Controls.Add(this.TbUniqueValues);
            this.Controls.Add(this.ClbColumns);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "UniqueDataForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Unique";
            this.Load += new System.EventHandler(this.UniqueDataForm_Load);
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
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button UniqueForm_Ok;
    }
}