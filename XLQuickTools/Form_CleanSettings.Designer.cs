namespace XLQuickTools
{
    partial class CleanSettingsForm
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
            this.TrimCleanForm_Cancel = new System.Windows.Forms.Button();
            this.TrimCleanForm_Save = new System.Windows.Forms.Button();
            this.CbNonPrintable = new System.Windows.Forms.CheckBox();
            this.CbSpaces = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.CbSpacesAll = new System.Windows.Forms.CheckBox();
            this.CbNonASCII = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TrimCleanForm_Cancel
            // 
            this.TrimCleanForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TrimCleanForm_Cancel.Location = new System.Drawing.Point(398, 280);
            this.TrimCleanForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.TrimCleanForm_Cancel.Name = "TrimCleanForm_Cancel";
            this.TrimCleanForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.TrimCleanForm_Cancel.TabIndex = 7;
            this.TrimCleanForm_Cancel.Text = "Cancel";
            this.TrimCleanForm_Cancel.UseVisualStyleBackColor = true;
            this.TrimCleanForm_Cancel.Click += new System.EventHandler(this.TrimCleanForm_Cancel_Click);
            // 
            // TrimCleanForm_Save
            // 
            this.TrimCleanForm_Save.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TrimCleanForm_Save.Location = new System.Drawing.Point(210, 280);
            this.TrimCleanForm_Save.Margin = new System.Windows.Forms.Padding(4);
            this.TrimCleanForm_Save.Name = "TrimCleanForm_Save";
            this.TrimCleanForm_Save.Size = new System.Drawing.Size(164, 56);
            this.TrimCleanForm_Save.TabIndex = 6;
            this.TrimCleanForm_Save.Text = "Save";
            this.TrimCleanForm_Save.UseVisualStyleBackColor = true;
            this.TrimCleanForm_Save.Click += new System.EventHandler(this.TrimCleanForm_Save_Click);
            // 
            // CbNonPrintable
            // 
            this.CbNonPrintable.AutoSize = true;
            this.CbNonPrintable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbNonPrintable.Location = new System.Drawing.Point(39, 128);
            this.CbNonPrintable.Margin = new System.Windows.Forms.Padding(6);
            this.CbNonPrintable.Name = "CbNonPrintable";
            this.CbNonPrintable.Size = new System.Drawing.Size(412, 33);
            this.CbNonPrintable.TabIndex = 3;
            this.CbNonPrintable.Text = "Remove Non-&Printable Characters";
            this.CbNonPrintable.UseVisualStyleBackColor = true;
            // 
            // CbSpaces
            // 
            this.CbSpaces.AutoSize = true;
            this.CbSpaces.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbSpaces.Location = new System.Drawing.Point(39, 48);
            this.CbSpaces.Margin = new System.Windows.Forms.Padding(6);
            this.CbSpaces.Name = "CbSpaces";
            this.CbSpaces.Size = new System.Drawing.Size(467, 33);
            this.CbSpaces.TabIndex = 1;
            this.CbSpaces.Text = "Remove &Two or More Spaces Between";
            this.CbSpaces.UseVisualStyleBackColor = true;
            this.CbSpaces.CheckedChanged += new System.EventHandler(this.CbSpaces_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CbSpaces);
            this.groupBox1.Controls.Add(this.CbSpacesAll);
            this.groupBox1.Controls.Add(this.CbNonPrintable);
            this.groupBox1.Controls.Add(this.CbNonASCII);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(22, 23);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(540, 237);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Options";
            // 
            // CbSpacesAll
            // 
            this.CbSpacesAll.AutoSize = true;
            this.CbSpacesAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbSpacesAll.Location = new System.Drawing.Point(39, 89);
            this.CbSpacesAll.Margin = new System.Windows.Forms.Padding(6);
            this.CbSpacesAll.Name = "CbSpacesAll";
            this.CbSpacesAll.Size = new System.Drawing.Size(356, 33);
            this.CbSpacesAll.TabIndex = 2;
            this.CbSpacesAll.Text = "Remove &All Spaces Between";
            this.CbSpacesAll.UseVisualStyleBackColor = true;
            this.CbSpacesAll.CheckedChanged += new System.EventHandler(this.CbSpacesAll_CheckedChanged);
            // 
            // CbNonASCII
            // 
            this.CbNonASCII.AutoSize = true;
            this.CbNonASCII.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CbNonASCII.Location = new System.Drawing.Point(39, 168);
            this.CbNonASCII.Margin = new System.Windows.Forms.Padding(6);
            this.CbNonASCII.Name = "CbNonASCII";
            this.CbNonASCII.Size = new System.Drawing.Size(376, 33);
            this.CbNonASCII.TabIndex = 4;
            this.CbNonASCII.Text = "Remove Non-A&SCII Characters";
            this.CbNonASCII.UseVisualStyleBackColor = true;
            // 
            // CleanSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(585, 356);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.TrimCleanForm_Cancel);
            this.Controls.Add(this.TrimCleanForm_Save);
            this.Name = "CleanSettingsForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Clean Settings";
            this.Load += new System.EventHandler(this.Form_TrimCleanSettings_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button TrimCleanForm_Cancel;
        private System.Windows.Forms.Button TrimCleanForm_Save;
        private System.Windows.Forms.CheckBox CbNonPrintable;
        private System.Windows.Forms.CheckBox CbSpaces;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox CbSpacesAll;
        private System.Windows.Forms.CheckBox CbNonASCII;
    }
}