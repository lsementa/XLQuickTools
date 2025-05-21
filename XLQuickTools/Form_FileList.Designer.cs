namespace XLQuickTools
{
    partial class FileListForm
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
            this.TbFolder = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.FileListForm_Cancel = new System.Windows.Forms.Button();
            this.FileListForm_Ok = new System.Windows.Forms.Button();
            this.FileListForm_Browse = new System.Windows.Forms.Button();
            this.TbSheetName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.CbChildFolders = new System.Windows.Forms.CheckBox();
            this.CbActiveSheet = new System.Windows.Forms.CheckBox();
            this.CbDateModified = new System.Windows.Forms.CheckBox();
            this.CbFileSize = new System.Windows.Forms.CheckBox();
            this.CbFileLocation = new System.Windows.Forms.CheckBox();
            this.CbFileType = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // TbFolder
            // 
            this.TbFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbFolder.Location = new System.Drawing.Point(29, 54);
            this.TbFolder.Margin = new System.Windows.Forms.Padding(6);
            this.TbFolder.Name = "TbFolder";
            this.TbFolder.Size = new System.Drawing.Size(596, 35);
            this.TbFolder.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(24, 19);
            this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(187, 29);
            this.label4.TabIndex = 0;
            this.label4.Text = "&Folder Location:";
            // 
            // FileListForm_Cancel
            // 
            this.FileListForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FileListForm_Cancel.Location = new System.Drawing.Point(569, 339);
            this.FileListForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.FileListForm_Cancel.Name = "FileListForm_Cancel";
            this.FileListForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.FileListForm_Cancel.TabIndex = 14;
            this.FileListForm_Cancel.Text = "Cancel";
            this.FileListForm_Cancel.UseVisualStyleBackColor = true;
            this.FileListForm_Cancel.Click += new System.EventHandler(this.FileListForm_Cancel_Click);
            // 
            // FileListForm_Ok
            // 
            this.FileListForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FileListForm_Ok.Location = new System.Drawing.Point(381, 339);
            this.FileListForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.FileListForm_Ok.Name = "FileListForm_Ok";
            this.FileListForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.FileListForm_Ok.TabIndex = 13;
            this.FileListForm_Ok.Text = "Ok";
            this.FileListForm_Ok.UseVisualStyleBackColor = true;
            this.FileListForm_Ok.Click += new System.EventHandler(this.FileListForm_Ok_Click);
            // 
            // FileListForm_Browse
            // 
            this.FileListForm_Browse.Location = new System.Drawing.Point(634, 51);
            this.FileListForm_Browse.Name = "FileListForm_Browse";
            this.FileListForm_Browse.Size = new System.Drawing.Size(99, 43);
            this.FileListForm_Browse.TabIndex = 2;
            this.FileListForm_Browse.Text = "Browse";
            this.FileListForm_Browse.UseVisualStyleBackColor = true;
            this.FileListForm_Browse.Click += new System.EventHandler(this.FileListForm_Browse_Click);
            // 
            // TbSheetName
            // 
            this.TbSheetName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TbSheetName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TbSheetName.Location = new System.Drawing.Point(25, 149);
            this.TbSheetName.Margin = new System.Windows.Forms.Padding(6);
            this.TbSheetName.Name = "TbSheetName";
            this.TbSheetName.Size = new System.Drawing.Size(412, 35);
            this.TbSheetName.TabIndex = 12;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(20, 114);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(190, 29);
            this.label1.TabIndex = 11;
            this.label1.Text = "&New Worksheet:";
            // 
            // CbChildFolders
            // 
            this.CbChildFolders.AutoSize = true;
            this.CbChildFolders.Location = new System.Drawing.Point(25, 42);
            this.CbChildFolders.Margin = new System.Windows.Forms.Padding(4);
            this.CbChildFolders.Name = "CbChildFolders";
            this.CbChildFolders.Size = new System.Drawing.Size(257, 29);
            this.CbChildFolders.TabIndex = 9;
            this.CbChildFolders.Text = "&Include Subdirectories";
            this.CbChildFolders.UseVisualStyleBackColor = true;
            // 
            // CbActiveSheet
            // 
            this.CbActiveSheet.AutoSize = true;
            this.CbActiveSheet.Location = new System.Drawing.Point(25, 79);
            this.CbActiveSheet.Margin = new System.Windows.Forms.Padding(4);
            this.CbActiveSheet.Name = "CbActiveSheet";
            this.CbActiveSheet.Size = new System.Drawing.Size(256, 29);
            this.CbActiveSheet.TabIndex = 10;
            this.CbActiveSheet.Text = "Use &Active Worksheet";
            this.CbActiveSheet.UseVisualStyleBackColor = true;
            this.CbActiveSheet.CheckedChanged += new System.EventHandler(this.CbActiveSheet_CheckedChanged);
            // 
            // CbDateModified
            // 
            this.CbDateModified.AutoSize = true;
            this.CbDateModified.Location = new System.Drawing.Point(25, 78);
            this.CbDateModified.Margin = new System.Windows.Forms.Padding(4);
            this.CbDateModified.Name = "CbDateModified";
            this.CbDateModified.Size = new System.Drawing.Size(177, 29);
            this.CbDateModified.TabIndex = 5;
            this.CbDateModified.Text = "&Date Modified";
            this.CbDateModified.UseVisualStyleBackColor = true;
            // 
            // CbFileSize
            // 
            this.CbFileSize.AutoSize = true;
            this.CbFileSize.Location = new System.Drawing.Point(25, 154);
            this.CbFileSize.Margin = new System.Windows.Forms.Padding(4);
            this.CbFileSize.Name = "CbFileSize";
            this.CbFileSize.Size = new System.Drawing.Size(127, 29);
            this.CbFileSize.TabIndex = 7;
            this.CbFileSize.Text = "File &Size";
            this.CbFileSize.UseVisualStyleBackColor = true;
            // 
            // CbFileLocation
            // 
            this.CbFileLocation.AutoSize = true;
            this.CbFileLocation.Location = new System.Drawing.Point(25, 42);
            this.CbFileLocation.Margin = new System.Windows.Forms.Padding(4);
            this.CbFileLocation.Name = "CbFileLocation";
            this.CbFileLocation.Size = new System.Drawing.Size(167, 29);
            this.CbFileLocation.TabIndex = 4;
            this.CbFileLocation.Text = "File &Location";
            this.CbFileLocation.UseVisualStyleBackColor = true;
            // 
            // CbFileType
            // 
            this.CbFileType.AutoSize = true;
            this.CbFileType.Location = new System.Drawing.Point(25, 115);
            this.CbFileType.Margin = new System.Windows.Forms.Padding(4);
            this.CbFileType.Name = "CbFileType";
            this.CbFileType.Size = new System.Drawing.Size(133, 29);
            this.CbFileType.TabIndex = 6;
            this.CbFileType.Text = "File &Type";
            this.CbFileType.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CbFileLocation);
            this.groupBox1.Controls.Add(this.CbFileType);
            this.groupBox1.Controls.Add(this.CbDateModified);
            this.groupBox1.Controls.Add(this.CbFileSize);
            this.groupBox1.Location = new System.Drawing.Point(29, 112);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(218, 208);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Display";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.CbChildFolders);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.CbActiveSheet);
            this.groupBox2.Controls.Add(this.TbSheetName);
            this.groupBox2.Location = new System.Drawing.Point(272, 112);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(461, 208);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Options";
            // 
            // FileListForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(763, 422);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.FileListForm_Browse);
            this.Controls.Add(this.FileListForm_Cancel);
            this.Controls.Add(this.FileListForm_Ok);
            this.Controls.Add(this.TbFolder);
            this.Controls.Add(this.label4);
            this.Name = "FileListForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Create File List";
            this.Load += new System.EventHandler(this.FileListForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TbFolder;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button FileListForm_Cancel;
        private System.Windows.Forms.Button FileListForm_Ok;
        private System.Windows.Forms.Button FileListForm_Browse;
        private System.Windows.Forms.TextBox TbSheetName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox CbChildFolders;
        private System.Windows.Forms.CheckBox CbActiveSheet;
        private System.Windows.Forms.CheckBox CbDateModified;
        private System.Windows.Forms.CheckBox CbFileSize;
        private System.Windows.Forms.CheckBox CbFileLocation;
        private System.Windows.Forms.CheckBox CbFileType;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}