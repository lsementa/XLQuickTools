namespace XLQuickTools
{
    partial class SplitterForm
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
            this.label2 = new System.Windows.Forms.Label();
            this.SplitterForm_Cancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SplitterForm_Ok = new System.Windows.Forms.Button();
            this.CbDelimiter = new System.Windows.Forms.ComboBox();
            this.CbHeaders = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // TbCustom
            // 
            this.TbCustom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbCustom.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbCustom.Location = new System.Drawing.Point(32, 208);
            this.TbCustom.Margin = new System.Windows.Forms.Padding(4);
            this.TbCustom.Name = "TbCustom";
            this.TbCustom.Size = new System.Drawing.Size(383, 35);
            this.TbCustom.TabIndex = 3;
            this.TbCustom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(27, 175);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 29);
            this.label2.TabIndex = 2;
            this.label2.Text = "&Custom:";
            // 
            // SplitterForm_Cancel
            // 
            this.SplitterForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SplitterForm_Cancel.Location = new System.Drawing.Point(251, 280);
            this.SplitterForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.SplitterForm_Cancel.Name = "SplitterForm_Cancel";
            this.SplitterForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.SplitterForm_Cancel.TabIndex = 5;
            this.SplitterForm_Cancel.Text = "Cancel";
            this.SplitterForm_Cancel.UseVisualStyleBackColor = true;
            this.SplitterForm_Cancel.Click += new System.EventHandler(this.SplitterForm_Cancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(27, 83);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "&Delimiter:";
            // 
            // SplitterForm_Ok
            // 
            this.SplitterForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SplitterForm_Ok.Location = new System.Drawing.Point(68, 280);
            this.SplitterForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.SplitterForm_Ok.Name = "SplitterForm_Ok";
            this.SplitterForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.SplitterForm_Ok.TabIndex = 4;
            this.SplitterForm_Ok.Text = "Ok";
            this.SplitterForm_Ok.UseVisualStyleBackColor = true;
            this.SplitterForm_Ok.Click += new System.EventHandler(this.SplitterForm_Ok_Click);
            // 
            // CbDelimiter
            // 
            this.CbDelimiter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbDelimiter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbDelimiter.FormattingEnabled = true;
            this.CbDelimiter.Location = new System.Drawing.Point(32, 116);
            this.CbDelimiter.Margin = new System.Windows.Forms.Padding(4);
            this.CbDelimiter.Name = "CbDelimiter";
            this.CbDelimiter.Size = new System.Drawing.Size(383, 37);
            this.CbDelimiter.TabIndex = 1;
            this.CbDelimiter.SelectedIndexChanged += new System.EventHandler(this.cbDelimiter_SelectedIndexChanged);
            // 
            // CbHeaders
            // 
            this.CbHeaders.AutoSize = true;
            this.CbHeaders.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbHeaders.Location = new System.Drawing.Point(32, 28);
            this.CbHeaders.Name = "CbHeaders";
            this.CbHeaders.Size = new System.Drawing.Size(266, 33);
            this.CbHeaders.TabIndex = 6;
            this.CbHeaders.Text = "&My data has headers";
            this.CbHeaders.UseVisualStyleBackColor = true;
            this.CbHeaders.CheckedChanged += new System.EventHandler(this.CbHeaders_CheckedChanged);
            // 
            // SplitterForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(449, 361);
            this.Controls.Add(this.CbHeaders);
            this.Controls.Add(this.CbDelimiter);
            this.Controls.Add(this.TbCustom);
            this.Controls.Add(this.SplitterForm_Cancel);
            this.Controls.Add(this.SplitterForm_Ok);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "SplitterForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Split to Rows";
            this.Load += new System.EventHandler(this.SplitterForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TbCustom;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button SplitterForm_Cancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button SplitterForm_Ok;
        private System.Windows.Forms.ComboBox CbDelimiter;
        private System.Windows.Forms.CheckBox CbHeaders;
    }
}