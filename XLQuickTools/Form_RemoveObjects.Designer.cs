namespace XLQuickTools
{
    partial class RemoveObjectsForm
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
            this.RemoveObjectsForm_Cancel = new System.Windows.Forms.Button();
            this.RemoveObjectsForm_Ok = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.CbShapes = new System.Windows.Forms.CheckBox();
            this.CbCharts = new System.Windows.Forms.CheckBox();
            this.CbActiveX = new System.Windows.Forms.CheckBox();
            this.CbFormControls = new System.Windows.Forms.CheckBox();
            this.CbComments = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // RemoveObjectsForm_Cancel
            // 
            this.RemoveObjectsForm_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RemoveObjectsForm_Cancel.Location = new System.Drawing.Point(336, 266);
            this.RemoveObjectsForm_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.RemoveObjectsForm_Cancel.Name = "RemoveObjectsForm_Cancel";
            this.RemoveObjectsForm_Cancel.Size = new System.Drawing.Size(164, 56);
            this.RemoveObjectsForm_Cancel.TabIndex = 15;
            this.RemoveObjectsForm_Cancel.Text = "Cancel";
            this.RemoveObjectsForm_Cancel.UseVisualStyleBackColor = true;
            this.RemoveObjectsForm_Cancel.Click += new System.EventHandler(this.RemoveObjectsForm_Cancel_Click);
            // 
            // RemoveObjectsForm_Ok
            // 
            this.RemoveObjectsForm_Ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RemoveObjectsForm_Ok.Location = new System.Drawing.Point(153, 266);
            this.RemoveObjectsForm_Ok.Margin = new System.Windows.Forms.Padding(4);
            this.RemoveObjectsForm_Ok.Name = "RemoveObjectsForm_Ok";
            this.RemoveObjectsForm_Ok.Size = new System.Drawing.Size(164, 56);
            this.RemoveObjectsForm_Ok.TabIndex = 14;
            this.RemoveObjectsForm_Ok.Text = "Ok";
            this.RemoveObjectsForm_Ok.UseVisualStyleBackColor = true;
            this.RemoveObjectsForm_Ok.Click += new System.EventHandler(this.RemoveObjectsForm_Ok_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CbShapes);
            this.groupBox1.Controls.Add(this.CbCharts);
            this.groupBox1.Controls.Add(this.CbActiveX);
            this.groupBox1.Controls.Add(this.CbFormControls);
            this.groupBox1.Controls.Add(this.CbComments);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(24, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(476, 234);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Objects";
            // 
            // CbShapes
            // 
            this.CbShapes.AutoSize = true;
            this.CbShapes.Location = new System.Drawing.Point(39, 42);
            this.CbShapes.Margin = new System.Windows.Forms.Padding(4);
            this.CbShapes.Name = "CbShapes";
            this.CbShapes.Size = new System.Drawing.Size(409, 33);
            this.CbShapes.TabIndex = 1;
            this.CbShapes.Text = "&Shapes (Pictures, SmartArt, Icons)";
            this.CbShapes.UseVisualStyleBackColor = true;
            // 
            // CbCharts
            // 
            this.CbCharts.AutoSize = true;
            this.CbCharts.Location = new System.Drawing.Point(39, 78);
            this.CbCharts.Margin = new System.Windows.Forms.Padding(4);
            this.CbCharts.Name = "CbCharts";
            this.CbCharts.Size = new System.Drawing.Size(114, 33);
            this.CbCharts.TabIndex = 2;
            this.CbCharts.Text = "&Charts";
            this.CbCharts.UseVisualStyleBackColor = true;
            // 
            // CbActiveX
            // 
            this.CbActiveX.AutoSize = true;
            this.CbActiveX.Location = new System.Drawing.Point(39, 115);
            this.CbActiveX.Margin = new System.Windows.Forms.Padding(4);
            this.CbActiveX.Name = "CbActiveX";
            this.CbActiveX.Size = new System.Drawing.Size(372, 33);
            this.CbActiveX.TabIndex = 3;
            this.CbActiveX.Text = "&OLEObjects / ActiveX Controls";
            this.CbActiveX.UseVisualStyleBackColor = true;
            // 
            // CbFormControls
            // 
            this.CbFormControls.AutoSize = true;
            this.CbFormControls.Location = new System.Drawing.Point(39, 151);
            this.CbFormControls.Margin = new System.Windows.Forms.Padding(4);
            this.CbFormControls.Name = "CbFormControls";
            this.CbFormControls.Size = new System.Drawing.Size(198, 33);
            this.CbFormControls.TabIndex = 4;
            this.CbFormControls.Text = "&Form Controls";
            this.CbFormControls.UseVisualStyleBackColor = true;
            // 
            // CbComments
            // 
            this.CbComments.AutoSize = true;
            this.CbComments.Location = new System.Drawing.Point(39, 188);
            this.CbComments.Margin = new System.Windows.Forms.Padding(4);
            this.CbComments.Name = "CbComments";
            this.CbComments.Size = new System.Drawing.Size(161, 33);
            this.CbComments.TabIndex = 5;
            this.CbComments.Text = "Co&mments";
            this.CbComments.UseVisualStyleBackColor = true;
            // 
            // RemoveObjectsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(526, 346);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.RemoveObjectsForm_Cancel);
            this.Controls.Add(this.RemoveObjectsForm_Ok);
            this.Name = "RemoveObjectsForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Remove Objects";
            this.Load += new System.EventHandler(this.RemoveObjectsForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button RemoveObjectsForm_Cancel;
        private System.Windows.Forms.Button RemoveObjectsForm_Ok;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox CbFormControls;
        private System.Windows.Forms.CheckBox CbActiveX;
        private System.Windows.Forms.CheckBox CbCharts;
        private System.Windows.Forms.CheckBox CbShapes;
        private System.Windows.Forms.CheckBox CbComments;
    }
}