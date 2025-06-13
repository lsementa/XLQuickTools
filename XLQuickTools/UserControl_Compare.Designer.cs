namespace XLQuickTools
{
    partial class UserControl_Compare
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.TbMaxRows = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.CbWorkbooks1 = new System.Windows.Forms.ComboBox();
            this.CbWorksheets1 = new System.Windows.Forms.ComboBox();
            this.CbWorkbooks2 = new System.Windows.Forms.ComboBox();
            this.CbWorksheets2 = new System.Windows.Forms.ComboBox();
            this.BtnClose = new System.Windows.Forms.Button();
            this.DgCompare = new System.Windows.Forms.DataGridView();
            this.BtnRun = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.BtnClear = new System.Windows.Forms.Button();
            this.CbHighlight = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.DgCompare)).BeginInit();
            this.SuspendLayout();
            // 
            // TbMaxRows
            // 
            this.TbMaxRows.Location = new System.Drawing.Point(508, 259);
            this.TbMaxRows.Margin = new System.Windows.Forms.Padding(6);
            this.TbMaxRows.Name = "TbMaxRows";
            this.TbMaxRows.Size = new System.Drawing.Size(86, 31);
            this.TbMaxRows.TabIndex = 4;
            this.TbMaxRows.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 263);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(479, 25);
            this.label1.TabIndex = 13;
            this.label1.Text = "Place results on a new worksheet if they exceed:";
            // 
            // CbWorkbooks1
            // 
            this.CbWorkbooks1.FormattingEnabled = true;
            this.CbWorkbooks1.Location = new System.Drawing.Point(26, 67);
            this.CbWorkbooks1.Margin = new System.Windows.Forms.Padding(6);
            this.CbWorkbooks1.Name = "CbWorkbooks1";
            this.CbWorkbooks1.Size = new System.Drawing.Size(243, 33);
            this.CbWorkbooks1.TabIndex = 0;
            // 
            // CbWorksheets1
            // 
            this.CbWorksheets1.FormattingEnabled = true;
            this.CbWorksheets1.Location = new System.Drawing.Point(26, 154);
            this.CbWorksheets1.Margin = new System.Windows.Forms.Padding(6);
            this.CbWorksheets1.Name = "CbWorksheets1";
            this.CbWorksheets1.Size = new System.Drawing.Size(243, 33);
            this.CbWorksheets1.TabIndex = 1;
            // 
            // CbWorkbooks2
            // 
            this.CbWorkbooks2.FormattingEnabled = true;
            this.CbWorkbooks2.Location = new System.Drawing.Point(301, 67);
            this.CbWorkbooks2.Margin = new System.Windows.Forms.Padding(6);
            this.CbWorkbooks2.Name = "CbWorkbooks2";
            this.CbWorkbooks2.Size = new System.Drawing.Size(246, 33);
            this.CbWorkbooks2.TabIndex = 2;
            // 
            // CbWorksheets2
            // 
            this.CbWorksheets2.FormattingEnabled = true;
            this.CbWorksheets2.Location = new System.Drawing.Point(301, 154);
            this.CbWorksheets2.Margin = new System.Windows.Forms.Padding(6);
            this.CbWorksheets2.Name = "CbWorksheets2";
            this.CbWorksheets2.Size = new System.Drawing.Size(246, 33);
            this.CbWorksheets2.TabIndex = 3;
            // 
            // BtnClose
            // 
            this.BtnClose.Location = new System.Drawing.Point(414, 749);
            this.BtnClose.Margin = new System.Windows.Forms.Padding(6);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(164, 56);
            this.BtnClose.TabIndex = 8;
            this.BtnClose.Text = "Close";
            this.BtnClose.UseVisualStyleBackColor = true;
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // DgCompare
            // 
            this.DgCompare.AllowUserToAddRows = false;
            this.DgCompare.AllowUserToDeleteRows = false;
            this.DgCompare.AllowUserToResizeColumns = false;
            this.DgCompare.AllowUserToResizeRows = false;
            this.DgCompare.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DgCompare.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DgCompare.BackgroundColor = System.Drawing.SystemColors.Window;
            this.DgCompare.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.DgCompare.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgCompare.EnableHeadersVisualStyles = false;
            this.DgCompare.Location = new System.Drawing.Point(26, 309);
            this.DgCompare.Margin = new System.Windows.Forms.Padding(6);
            this.DgCompare.Name = "DgCompare";
            this.DgCompare.ReadOnly = true;
            this.DgCompare.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.DgCompare.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.DgCompare.ShowCellErrors = false;
            this.DgCompare.ShowCellToolTips = false;
            this.DgCompare.ShowEditingIcon = false;
            this.DgCompare.ShowRowErrors = false;
            this.DgCompare.Size = new System.Drawing.Size(828, 429);
            this.DgCompare.TabIndex = 5;
            // 
            // BtnRun
            // 
            this.BtnRun.Location = new System.Drawing.Point(26, 749);
            this.BtnRun.Margin = new System.Windows.Forms.Padding(6);
            this.BtnRun.Name = "BtnRun";
            this.BtnRun.Size = new System.Drawing.Size(164, 56);
            this.BtnRun.TabIndex = 6;
            this.BtnRun.Text = "Run";
            this.BtnRun.UseVisualStyleBackColor = true;
            this.BtnRun.Click += new System.EventHandler(this.BtnRun_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label2.Location = new System.Drawing.Point(28, 33);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(148, 29);
            this.label2.TabIndex = 23;
            this.label2.Text = "Workbook 1:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label3.Location = new System.Drawing.Point(296, 33);
            this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(148, 29);
            this.label3.TabIndex = 24;
            this.label3.Text = "Workbook 2:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label4.Location = new System.Drawing.Point(22, 119);
            this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(153, 29);
            this.label4.TabIndex = 25;
            this.label4.Text = "Worksheet 1:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label5.Location = new System.Drawing.Point(296, 119);
            this.label5.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(153, 29);
            this.label5.TabIndex = 26;
            this.label5.Text = "Worksheet 2:";
            // 
            // BtnClear
            // 
            this.BtnClear.Location = new System.Drawing.Point(221, 749);
            this.BtnClear.Margin = new System.Windows.Forms.Padding(6);
            this.BtnClear.Name = "BtnClear";
            this.BtnClear.Size = new System.Drawing.Size(164, 56);
            this.BtnClear.TabIndex = 7;
            this.BtnClear.Text = "Refresh";
            this.BtnClear.UseVisualStyleBackColor = true;
            this.BtnClear.Click += new System.EventHandler(this.BtnClear_Click);
            // 
            // CbHighlight
            // 
            this.CbHighlight.AutoSize = true;
            this.CbHighlight.Location = new System.Drawing.Point(26, 214);
            this.CbHighlight.Name = "CbHighlight";
            this.CbHighlight.Size = new System.Drawing.Size(374, 29);
            this.CbHighlight.TabIndex = 27;
            this.CbHighlight.Text = "Highlight differences on worksheet";
            this.CbHighlight.UseVisualStyleBackColor = true;
            // 
            // UserControl_Compare
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.Controls.Add(this.CbHighlight);
            this.Controls.Add(this.BtnClear);
            this.Controls.Add(this.BtnRun);
            this.Controls.Add(this.DgCompare);
            this.Controls.Add(this.BtnClose);
            this.Controls.Add(this.CbWorksheets2);
            this.Controls.Add(this.CbWorkbooks2);
            this.Controls.Add(this.CbWorksheets1);
            this.Controls.Add(this.CbWorkbooks1);
            this.Controls.Add(this.TbMaxRows);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "UserControl_Compare";
            this.Size = new System.Drawing.Size(892, 825);
            this.Load += new System.EventHandler(this.UserControl_Compare_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DgCompare)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox TbMaxRows;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox CbWorkbooks1;
        private System.Windows.Forms.ComboBox CbWorksheets1;
        private System.Windows.Forms.ComboBox CbWorkbooks2;
        private System.Windows.Forms.ComboBox CbWorksheets2;
        private System.Windows.Forms.Button BtnClose;
        private System.Windows.Forms.DataGridView DgCompare;
        private System.Windows.Forms.Button BtnRun;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button BtnClear;
        private System.Windows.Forms.CheckBox CbHighlight;
    }
}
