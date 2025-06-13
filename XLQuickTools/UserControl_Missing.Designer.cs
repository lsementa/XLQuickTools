namespace XLQuickTools
{
    partial class UserControl_Missing
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
            this.TbRange1 = new System.Windows.Forms.TextBox();
            this.TbRange2 = new System.Windows.Forms.TextBox();
            this.DgMissing = new System.Windows.Forms.DataGridView();
            this.UcMissingClose = new System.Windows.Forms.Button();
            this.BtnCheck = new System.Windows.Forms.Button();
            this.TbMaxRows = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.BtnClear = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DgMissing)).BeginInit();
            this.SuspendLayout();
            // 
            // TbRange1
            // 
            this.TbRange1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.TbRange1.Location = new System.Drawing.Point(32, 65);
            this.TbRange1.Margin = new System.Windows.Forms.Padding(6);
            this.TbRange1.Name = "TbRange1";
            this.TbRange1.Size = new System.Drawing.Size(568, 31);
            this.TbRange1.TabIndex = 1;
            this.TbRange1.Click += new System.EventHandler(this.TbRange1_Click);
            // 
            // TbRange2
            // 
            this.TbRange2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.TbRange2.Location = new System.Drawing.Point(32, 156);
            this.TbRange2.Margin = new System.Windows.Forms.Padding(6);
            this.TbRange2.Name = "TbRange2";
            this.TbRange2.Size = new System.Drawing.Size(568, 31);
            this.TbRange2.TabIndex = 3;
            this.TbRange2.Click += new System.EventHandler(this.TbRange2_Click);
            this.TbRange2.TextChanged += new System.EventHandler(this.TbRange2_TextChanged);
            // 
            // DgMissing
            // 
            this.DgMissing.AllowUserToAddRows = false;
            this.DgMissing.AllowUserToDeleteRows = false;
            this.DgMissing.AllowUserToResizeColumns = false;
            this.DgMissing.AllowUserToResizeRows = false;
            this.DgMissing.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DgMissing.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DgMissing.BackgroundColor = System.Drawing.SystemColors.Window;
            this.DgMissing.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.DgMissing.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgMissing.EnableHeadersVisualStyles = false;
            this.DgMissing.Location = new System.Drawing.Point(32, 277);
            this.DgMissing.Margin = new System.Windows.Forms.Padding(6);
            this.DgMissing.Name = "DgMissing";
            this.DgMissing.ReadOnly = true;
            this.DgMissing.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.DgMissing.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.DgMissing.ShowCellErrors = false;
            this.DgMissing.ShowCellToolTips = false;
            this.DgMissing.ShowEditingIcon = false;
            this.DgMissing.ShowRowErrors = false;
            this.DgMissing.Size = new System.Drawing.Size(830, 429);
            this.DgMissing.TabIndex = 6;
            // 
            // UcMissingClose
            // 
            this.UcMissingClose.Location = new System.Drawing.Point(409, 715);
            this.UcMissingClose.Margin = new System.Windows.Forms.Padding(4);
            this.UcMissingClose.Name = "UcMissingClose";
            this.UcMissingClose.Size = new System.Drawing.Size(164, 56);
            this.UcMissingClose.TabIndex = 9;
            this.UcMissingClose.Text = "Close";
            this.UcMissingClose.UseVisualStyleBackColor = true;
            this.UcMissingClose.Click += new System.EventHandler(this.UcMissingClose_Click);
            // 
            // BtnCheck
            // 
            this.BtnCheck.Location = new System.Drawing.Point(32, 715);
            this.BtnCheck.Margin = new System.Windows.Forms.Padding(4);
            this.BtnCheck.Name = "BtnCheck";
            this.BtnCheck.Size = new System.Drawing.Size(164, 56);
            this.BtnCheck.TabIndex = 7;
            this.BtnCheck.Text = "Run";
            this.BtnCheck.UseVisualStyleBackColor = true;
            this.BtnCheck.Click += new System.EventHandler(this.BtnCheck_Click);
            // 
            // TbMaxRows
            // 
            this.TbMaxRows.Location = new System.Drawing.Point(514, 227);
            this.TbMaxRows.Margin = new System.Windows.Forms.Padding(6);
            this.TbMaxRows.Name = "TbMaxRows";
            this.TbMaxRows.Size = new System.Drawing.Size(86, 31);
            this.TbMaxRows.TabIndex = 5;
            this.TbMaxRows.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 233);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(479, 25);
            this.label1.TabIndex = 4;
            this.label1.Text = "Place results on a new worksheet if they exceed:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label2.Location = new System.Drawing.Point(26, 31);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 29);
            this.label2.TabIndex = 0;
            this.label2.Text = "Range 1:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label3.Location = new System.Drawing.Point(26, 121);
            this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(109, 29);
            this.label3.TabIndex = 2;
            this.label3.Text = "Range 2:";
            // 
            // BtnClear
            // 
            this.BtnClear.Location = new System.Drawing.Point(221, 715);
            this.BtnClear.Margin = new System.Windows.Forms.Padding(4);
            this.BtnClear.Name = "BtnClear";
            this.BtnClear.Size = new System.Drawing.Size(164, 56);
            this.BtnClear.TabIndex = 8;
            this.BtnClear.Text = "Refresh";
            this.BtnClear.UseVisualStyleBackColor = true;
            this.BtnClear.Click += new System.EventHandler(this.BtnClear_Click);
            // 
            // UserControl_Missing
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.Controls.Add(this.BtnClear);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.TbMaxRows);
            this.Controls.Add(this.BtnCheck);
            this.Controls.Add(this.UcMissingClose);
            this.Controls.Add(this.DgMissing);
            this.Controls.Add(this.TbRange1);
            this.Controls.Add(this.TbRange2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "UserControl_Missing";
            this.Size = new System.Drawing.Size(894, 792);
            this.Load += new System.EventHandler(this.UserControl_Missing_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DgMissing)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TbRange1;
        private System.Windows.Forms.TextBox TbRange2;
        private System.Windows.Forms.DataGridView DgMissing;
        private System.Windows.Forms.Button UcMissingClose;
        private System.Windows.Forms.Button BtnCheck;
        private System.Windows.Forms.TextBox TbMaxRows;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button BtnClear;
    }
}
