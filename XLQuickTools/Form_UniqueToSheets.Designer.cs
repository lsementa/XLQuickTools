namespace XLQuickTools
{
    partial class UniqueSheetsForm
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
            this.UniqueSheetsForm_Cancel = new System.Windows.Forms.Button();
            this.UniqueSheetsForm_Ok = new System.Windows.Forms.Button();
            this.TbUniqueValues = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // UniqueSheetsForm_Cancel
            // 
            this.UniqueSheetsForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UniqueSheetsForm_Cancel.Location = new System.Drawing.Point(413, 159);
            this.UniqueSheetsForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.UniqueSheetsForm_Cancel.Name = "UniqueSheetsForm_Cancel";
            this.UniqueSheetsForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.UniqueSheetsForm_Cancel.TabIndex = 13;
            this.UniqueSheetsForm_Cancel.Text = "Cancel";
            this.UniqueSheetsForm_Cancel.UseVisualStyleBackColor = true;
            this.UniqueSheetsForm_Cancel.Click += new System.EventHandler(this.UniqueSheetsForm_Cancel_Click);
            // 
            // UniqueSheetsForm_Ok
            // 
            this.UniqueSheetsForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UniqueSheetsForm_Ok.Location = new System.Drawing.Point(227, 159);
            this.UniqueSheetsForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.UniqueSheetsForm_Ok.Name = "UniqueSheetsForm_Ok";
            this.UniqueSheetsForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.UniqueSheetsForm_Ok.TabIndex = 12;
            this.UniqueSheetsForm_Ok.Text = "Ok";
            this.UniqueSheetsForm_Ok.UseVisualStyleBackColor = true;
            this.UniqueSheetsForm_Ok.Click += new System.EventHandler(this.UniqueSheetsForm_Ok_Click);
            // 
            // TbUniqueValues
            // 
            this.TbUniqueValues.BackColor = System.Drawing.Color.Gainsboro;
            this.TbUniqueValues.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbUniqueValues.Enabled = false;
            this.TbUniqueValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbUniqueValues.Location = new System.Drawing.Point(448, 49);
            this.TbUniqueValues.Name = "TbUniqueValues";
            this.TbUniqueValues.Size = new System.Drawing.Size(129, 35);
            this.TbUniqueValues.TabIndex = 15;
            this.TbUniqueValues.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(21, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(423, 29);
            this.label1.TabIndex = 14;
            this.label1.Text = "Number of Worksheets this will create:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(21, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(226, 29);
            this.label2.TabIndex = 16;
            this.label2.Text = "Click Ok to continue";
            // 
            // UniqueSheetsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(606, 245);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.TbUniqueValues);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.UniqueSheetsForm_Cancel);
            this.Controls.Add(this.UniqueSheetsForm_Ok);
            this.Name = "UniqueSheetsForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Create Worksheets";
            this.Load += new System.EventHandler(this.UniqueSheetsForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button UniqueSheetsForm_Cancel;
        private System.Windows.Forms.Button UniqueSheetsForm_Ok;
        private System.Windows.Forms.TextBox TbUniqueValues;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}