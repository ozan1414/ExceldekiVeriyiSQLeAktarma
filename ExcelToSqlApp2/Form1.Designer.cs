namespace ExcelToSqlApp2
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.OpenFileDialog openFileDialog;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();

            // Button (Excel Dosyası Seç)
            this.btnSelectFile.Location = new System.Drawing.Point(12, 12);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(150, 23);
            this.btnSelectFile.TabIndex = 0;
            this.btnSelectFile.Text = "Excel Dosyasını Seç";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);

            // TextBox (Log)
            this.txtLog.Location = new System.Drawing.Point(12, 50);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(360, 200);
            this.txtLog.TabIndex = 1;

            // Form
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 261);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.btnSelectFile);
            this.Name = "Form1";
            this.Text = "Excel to SQL Aktarım";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
