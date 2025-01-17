namespace XLQuickTools
{
    partial class LeadTrailForm
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
            this.TbTrailing = new System.Windows.Forms.TextBox();
            this.TbLeading = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.LeadTrailForm_Cancel = new System.Windows.Forms.Button();
            this.LeadTrailForm_Ok = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // TbTrailing
            // 
            this.TbTrailing.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbTrailing.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbTrailing.Location = new System.Drawing.Point(34, 148);
            this.TbTrailing.Margin = new System.Windows.Forms.Padding(4);
            this.TbTrailing.Name = "TbTrailing";
            this.TbTrailing.Size = new System.Drawing.Size(365, 35);
            this.TbTrailing.TabIndex = 3;
            this.TbTrailing.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // TbLeading
            // 
            this.TbLeading.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbLeading.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbLeading.Location = new System.Drawing.Point(34, 56);
            this.TbLeading.Margin = new System.Windows.Forms.Padding(4);
            this.TbLeading.Name = "TbLeading";
            this.TbLeading.Size = new System.Drawing.Size(365, 35);
            this.TbLeading.TabIndex = 1;
            this.TbLeading.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(29, 23);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(282, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "&Leading character or text:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(29, 115);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(277, 29);
            this.label2.TabIndex = 2;
            this.label2.Text = "&Trailing character or text:";
            // 
            // LeadTrailForm_Cancel
            // 
            this.LeadTrailForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LeadTrailForm_Cancel.Location = new System.Drawing.Point(225, 210);
            this.LeadTrailForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.LeadTrailForm_Cancel.Name = "LeadTrailForm_Cancel";
            this.LeadTrailForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.LeadTrailForm_Cancel.TabIndex = 5;
            this.LeadTrailForm_Cancel.Text = "Cancel";
            this.LeadTrailForm_Cancel.UseVisualStyleBackColor = true;
            this.LeadTrailForm_Cancel.Click += new System.EventHandler(this.LeadTrailForm_Cancel_Click);
            // 
            // LeadTrailForm_Ok
            // 
            this.LeadTrailForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LeadTrailForm_Ok.Location = new System.Drawing.Point(42, 210);
            this.LeadTrailForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.LeadTrailForm_Ok.Name = "LeadTrailForm_Ok";
            this.LeadTrailForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.LeadTrailForm_Ok.TabIndex = 4;
            this.LeadTrailForm_Ok.Text = "Ok";
            this.LeadTrailForm_Ok.UseVisualStyleBackColor = true;
            this.LeadTrailForm_Ok.Click += new System.EventHandler(this.LeadTrailForm_Ok_Click);
            // 
            // LeadTrailForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(430, 296);
            this.Controls.Add(this.LeadTrailForm_Cancel);
            this.Controls.Add(this.LeadTrailForm_Ok);
            this.Controls.Add(this.TbTrailing);
            this.Controls.Add(this.TbLeading);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Name = "LeadTrailForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Add Leading or Trailing";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TbTrailing;
        private System.Windows.Forms.TextBox TbLeading;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button LeadTrailForm_Cancel;
        private System.Windows.Forms.Button LeadTrailForm_Ok;
    }
}