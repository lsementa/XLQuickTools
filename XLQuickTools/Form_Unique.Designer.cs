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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UniqueDataForm));
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
            this.CbHeadersInclude = new System.Windows.Forms.CheckBox();
            this.CbHasDelimiter = new System.Windows.Forms.CheckBox();
            this.CbDelimiter = new System.Windows.Forms.ComboBox();
            this.TbCustom = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ClbColumns
            // 
            this.ClbColumns.CheckOnClick = true;
            this.ClbColumns.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ClbColumns.FormattingEnabled = true;
            this.ClbColumns.Location = new System.Drawing.Point(29, 441);
            this.ClbColumns.Name = "ClbColumns";
            this.ClbColumns.Size = new System.Drawing.Size(617, 228);
            this.ClbColumns.TabIndex = 9;
            this.ClbColumns.SelectedIndexChanged += new System.EventHandler(this.ClbColumns_SelectedIndexChanged);
            // 
            // CbHeaders
            // 
            this.CbHeaders.AutoSize = true;
            this.CbHeaders.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbHeaders.Location = new System.Drawing.Point(34, 151);
            this.CbHeaders.Name = "CbHeaders";
            this.CbHeaders.Size = new System.Drawing.Size(266, 33);
            this.CbHeaders.TabIndex = 5;
            this.CbHeaders.Text = "My data has &headers";
            this.CbHeaders.UseVisualStyleBackColor = true;
            this.CbHeaders.CheckedChanged += new System.EventHandler(this.Cb_Headers_CheckedChanged);
            // 
            // TbUniqueValues
            // 
            this.TbUniqueValues.BackColor = System.Drawing.Color.Gainsboro;
            this.TbUniqueValues.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbUniqueValues.Enabled = false;
            this.TbUniqueValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbUniqueValues.Location = new System.Drawing.Point(228, 32);
            this.TbUniqueValues.Name = "TbUniqueValues";
            this.TbUniqueValues.Size = new System.Drawing.Size(266, 35);
            this.TbUniqueValues.TabIndex = 3;
            this.TbUniqueValues.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(29, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(175, 29);
            this.label1.TabIndex = 1;
            this.label1.Text = "Unique Values:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(29, 92);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(163, 29);
            this.label2.TabIndex = 2;
            this.label2.Text = "Unique Rows:";
            // 
            // TbUniqueRows
            // 
            this.TbUniqueRows.BackColor = System.Drawing.Color.Gainsboro;
            this.TbUniqueRows.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbUniqueRows.Enabled = false;
            this.TbUniqueRows.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbUniqueRows.Location = new System.Drawing.Point(228, 92);
            this.TbUniqueRows.Name = "TbUniqueRows";
            this.TbUniqueRows.Size = new System.Drawing.Size(266, 35);
            this.TbUniqueRows.TabIndex = 4;
            this.TbUniqueRows.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // BtnSelectAll
            // 
            this.BtnSelectAll.Location = new System.Drawing.Point(348, 392);
            this.BtnSelectAll.Name = "BtnSelectAll";
            this.BtnSelectAll.Size = new System.Drawing.Size(146, 43);
            this.BtnSelectAll.TabIndex = 7;
            this.BtnSelectAll.Text = "Select &All";
            this.BtnSelectAll.UseVisualStyleBackColor = true;
            this.BtnSelectAll.Click += new System.EventHandler(this.BtnSelectAll_Click);
            // 
            // BtnUnselectAll
            // 
            this.BtnUnselectAll.Location = new System.Drawing.Point(500, 392);
            this.BtnUnselectAll.Name = "BtnUnselectAll";
            this.BtnUnselectAll.Size = new System.Drawing.Size(146, 43);
            this.BtnUnselectAll.TabIndex = 8;
            this.BtnUnselectAll.Text = "&Unselect All";
            this.BtnUnselectAll.UseVisualStyleBackColor = true;
            this.BtnUnselectAll.Click += new System.EventHandler(this.BtnUnselectAll_Click);
            // 
            // UniqueForm_Ok
            // 
            this.UniqueForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UniqueForm_Ok.Location = new System.Drawing.Point(296, 680);
            this.UniqueForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.UniqueForm_Ok.Name = "UniqueForm_Ok";
            this.UniqueForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.UniqueForm_Ok.TabIndex = 10;
            this.UniqueForm_Ok.Text = "Ok";
            this.UniqueForm_Ok.UseVisualStyleBackColor = true;
            this.UniqueForm_Ok.Click += new System.EventHandler(this.UniqueForm_Ok_Click);
            // 
            // UniqueForm_Cancel
            // 
            this.UniqueForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UniqueForm_Cancel.Location = new System.Drawing.Point(482, 680);
            this.UniqueForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.UniqueForm_Cancel.Name = "UniqueForm_Cancel";
            this.UniqueForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.UniqueForm_Cancel.TabIndex = 11;
            this.UniqueForm_Cancel.Text = "Cancel";
            this.UniqueForm_Cancel.UseVisualStyleBackColor = true;
            this.UniqueForm_Cancel.Click += new System.EventHandler(this.UniqueForm_Cancel_Click);
            // 
            // CbHeadersInclude
            // 
            this.CbHeadersInclude.AutoSize = true;
            this.CbHeadersInclude.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbHeadersInclude.Location = new System.Drawing.Point(34, 392);
            this.CbHeadersInclude.Name = "CbHeadersInclude";
            this.CbHeadersInclude.Size = new System.Drawing.Size(280, 33);
            this.CbHeadersInclude.TabIndex = 6;
            this.CbHeadersInclude.Text = "Add headers to &output";
            this.CbHeadersInclude.UseVisualStyleBackColor = true;
            // 
            // CbHasDelimiter
            // 
            this.CbHasDelimiter.AutoSize = true;
            this.CbHasDelimiter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbHasDelimiter.Location = new System.Drawing.Point(343, 151);
            this.CbHasDelimiter.Name = "CbHasDelimiter";
            this.CbHasDelimiter.Size = new System.Drawing.Size(142, 33);
            this.CbHasDelimiter.TabIndex = 12;
            this.CbHasDelimiter.Text = "&Delimiter";
            this.CbHasDelimiter.UseVisualStyleBackColor = true;
            this.CbHasDelimiter.CheckedChanged += new System.EventHandler(this.CbHasDelimiter_CheckedChanged);
            // 
            // CbDelimiter
            // 
            this.CbDelimiter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbDelimiter.Enabled = false;
            this.CbDelimiter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbDelimiter.FormattingEnabled = true;
            this.CbDelimiter.Location = new System.Drawing.Point(34, 236);
            this.CbDelimiter.Margin = new System.Windows.Forms.Padding(4);
            this.CbDelimiter.Name = "CbDelimiter";
            this.CbDelimiter.Size = new System.Drawing.Size(460, 37);
            this.CbDelimiter.TabIndex = 14;
            this.CbDelimiter.SelectedIndexChanged += new System.EventHandler(this.CbDelimiter_SelectedIndexChanged);
            // 
            // TbCustom
            // 
            this.TbCustom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbCustom.Enabled = false;
            this.TbCustom.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbCustom.Location = new System.Drawing.Point(34, 328);
            this.TbCustom.Margin = new System.Windows.Forms.Padding(4);
            this.TbCustom.Name = "TbCustom";
            this.TbCustom.Size = new System.Drawing.Size(230, 35);
            this.TbCustom.TabIndex = 16;
            this.TbCustom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TbCustom.TextChanged += new System.EventHandler(this.TbCustom_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(29, 203);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(116, 29);
            this.label3.TabIndex = 13;
            this.label3.Text = "De&limiter:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(29, 295);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(101, 29);
            this.label4.TabIndex = 15;
            this.label4.Text = "&Custom:";
            // 
            // UniqueDataForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(678, 757);
            this.Controls.Add(this.CbDelimiter);
            this.Controls.Add(this.TbCustom);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CbHasDelimiter);
            this.Controls.Add(this.TbUniqueValues);
            this.Controls.Add(this.TbUniqueRows);
            this.Controls.Add(this.UniqueForm_Cancel);
            this.Controls.Add(this.UniqueForm_Ok);
            this.Controls.Add(this.BtnUnselectAll);
            this.Controls.Add(this.BtnSelectAll);
            this.Controls.Add(this.ClbColumns);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.CbHeadersInclude);
            this.Controls.Add(this.CbHeaders);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "UniqueDataForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Selection to Clipboard";
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
        private System.Windows.Forms.Button UniqueForm_Ok;
        private System.Windows.Forms.Button UniqueForm_Cancel;
        private System.Windows.Forms.CheckBox CbHeadersInclude;
        private System.Windows.Forms.CheckBox CbHasDelimiter;
        private System.Windows.Forms.ComboBox CbDelimiter;
        private System.Windows.Forms.TextBox TbCustom;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}