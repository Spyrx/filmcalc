namespace MingoPackaging
{
    partial class exportForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(exportForm));
            this.btnCancel = new System.Windows.Forms.Button();
            this.labelCompanyName = new System.Windows.Forms.Label();
            this.labelBarName = new System.Windows.Forms.Label();
            this.labelBarFlavor = new System.Windows.Forms.Label();
            this.btnExport = new System.Windows.Forms.Button();
            this.txtCompanyName = new System.Windows.Forms.TextBox();
            this.txtBarName = new System.Windows.Forms.TextBox();
            this.txtBarFlavor = new System.Windows.Forms.TextBox();
            this.btnExit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(216, 127);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Back";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // labelCompanyName
            // 
            this.labelCompanyName.AutoSize = true;
            this.labelCompanyName.Location = new System.Drawing.Point(12, 16);
            this.labelCompanyName.Name = "labelCompanyName";
            this.labelCompanyName.Size = new System.Drawing.Size(82, 13);
            this.labelCompanyName.TabIndex = 1;
            this.labelCompanyName.Text = "Company Name";
            // 
            // labelBarName
            // 
            this.labelBarName.AutoSize = true;
            this.labelBarName.Location = new System.Drawing.Point(14, 43);
            this.labelBarName.Name = "labelBarName";
            this.labelBarName.Size = new System.Drawing.Size(54, 13);
            this.labelBarName.TabIndex = 2;
            this.labelBarName.Text = "Bar Name";
            // 
            // labelBarFlavor
            // 
            this.labelBarFlavor.AutoSize = true;
            this.labelBarFlavor.Location = new System.Drawing.Point(13, 73);
            this.labelBarFlavor.Name = "labelBarFlavor";
            this.labelBarFlavor.Size = new System.Drawing.Size(55, 13);
            this.labelBarFlavor.TabIndex = 3;
            this.labelBarFlavor.Text = "Bar Flavor";
            // 
            // btnExport
            // 
            this.btnExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExport.Location = new System.Drawing.Point(42, 115);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(116, 35);
            this.btnExport.TabIndex = 4;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.button2_Click);
            // 
            // txtCompanyName
            // 
            this.txtCompanyName.Location = new System.Drawing.Point(118, 13);
            this.txtCompanyName.Name = "txtCompanyName";
            this.txtCompanyName.Size = new System.Drawing.Size(254, 20);
            this.txtCompanyName.TabIndex = 5;
            // 
            // txtBarName
            // 
            this.txtBarName.Location = new System.Drawing.Point(118, 40);
            this.txtBarName.Name = "txtBarName";
            this.txtBarName.Size = new System.Drawing.Size(254, 20);
            this.txtBarName.TabIndex = 6;
            // 
            // txtBarFlavor
            // 
            this.txtBarFlavor.Location = new System.Drawing.Point(118, 70);
            this.txtBarFlavor.Name = "txtBarFlavor";
            this.txtBarFlavor.Size = new System.Drawing.Size(254, 20);
            this.txtBarFlavor.TabIndex = 7;
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(297, 127);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 8;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // exportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(402, 168);
            this.ControlBox = false;
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.txtBarFlavor);
            this.Controls.Add(this.txtBarName);
            this.Controls.Add(this.txtCompanyName);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.labelBarFlavor);
            this.Controls.Add(this.labelBarName);
            this.Controls.Add(this.labelCompanyName);
            this.Controls.Add(this.btnCancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "exportForm";
            this.Text = "Export";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label labelCompanyName;
        private System.Windows.Forms.Label labelBarName;
        private System.Windows.Forms.Label labelBarFlavor;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.TextBox txtCompanyName;
        private System.Windows.Forms.TextBox txtBarName;
        private System.Windows.Forms.TextBox txtBarFlavor;
        private System.Windows.Forms.Button btnExit;
    }
}