namespace XLQuickTools
{
    partial class HighlightForm
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
            this.HighlightForm_Cancel = new System.Windows.Forms.Button();
            this.HighlightForm_Ok = new System.Windows.Forms.Button();
            this.CbContainsHeaders = new System.Windows.Forms.CheckBox();
            this.TbSheetName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // HighlightForm_Cancel
            // 
            this.HighlightForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.HighlightForm_Cancel.Location = new System.Drawing.Point(315, 173);
            this.HighlightForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.HighlightForm_Cancel.Name = "HighlightForm_Cancel";
            this.HighlightForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.HighlightForm_Cancel.TabIndex = 4;
            this.HighlightForm_Cancel.Text = "Cancel";
            this.HighlightForm_Cancel.UseVisualStyleBackColor = true;
            this.HighlightForm_Cancel.Click += new System.EventHandler(this.HighlightForm_Cancel_Click);
            // 
            // HighlightForm_Ok
            // 
            this.HighlightForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.HighlightForm_Ok.Location = new System.Drawing.Point(132, 173);
            this.HighlightForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.HighlightForm_Ok.Name = "HighlightForm_Ok";
            this.HighlightForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.HighlightForm_Ok.TabIndex = 3;
            this.HighlightForm_Ok.Text = "Ok";
            this.HighlightForm_Ok.UseVisualStyleBackColor = true;
            this.HighlightForm_Ok.Click += new System.EventHandler(this.HighlightForm_Ok_Click);
            // 
            // CbContainsHeaders
            // 
            this.CbContainsHeaders.AutoSize = true;
            this.CbContainsHeaders.BackColor = System.Drawing.SystemColors.Control;
            this.CbContainsHeaders.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbContainsHeaders.Location = new System.Drawing.Point(32, 27);
            this.CbContainsHeaders.Name = "CbContainsHeaders";
            this.CbContainsHeaders.Size = new System.Drawing.Size(327, 33);
            this.CbContainsHeaders.TabIndex = 0;
            this.CbContainsHeaders.Text = "&First row contains headers";
            this.CbContainsHeaders.UseVisualStyleBackColor = false;
            // 
            // TbSheetName
            // 
            this.TbSheetName.BackColor = System.Drawing.SystemColors.Window;
            this.TbSheetName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbSheetName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbSheetName.Location = new System.Drawing.Point(32, 112);
            this.TbSheetName.Margin = new System.Windows.Forms.Padding(4);
            this.TbSheetName.Name = "TbSheetName";
            this.TbSheetName.Size = new System.Drawing.Size(447, 35);
            this.TbSheetName.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(27, 79);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(190, 29);
            this.label1.TabIndex = 1;
            this.label1.Text = "&New Worksheet:";
            // 
            // HighlightForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(513, 253);
            this.Controls.Add(this.TbSheetName);
            this.Controls.Add(this.CbContainsHeaders);
            this.Controls.Add(this.HighlightForm_Cancel);
            this.Controls.Add(this.HighlightForm_Ok);
            this.Controls.Add(this.label1);
            this.Name = "HighlightForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Copy Highlighted Rows";
            this.Load += new System.EventHandler(this.HighlightForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button HighlightForm_Cancel;
        private System.Windows.Forms.Button HighlightForm_Ok;
        private System.Windows.Forms.CheckBox CbContainsHeaders;
        private System.Windows.Forms.TextBox TbSheetName;
        private System.Windows.Forms.Label label1;
    }
}