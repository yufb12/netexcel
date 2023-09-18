 
namespace Feng.DataDesign
{
    partial class frmForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmForm));
            this.dataExcel1 = new Excel.DataExcelControl();
            this.SuspendLayout();
            // 
            // dataExcel1
            //  
            this.dataExcel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataExcel1.Font = new System.Drawing.Font("Tahoma", 9F); 
            this.dataExcel1.Location = new System.Drawing.Point(0, 0); 
            this.dataExcel1.Name = "dataExcel1"; 
            this.dataExcel1.Size = new System.Drawing.Size(722, 476);
            this.dataExcel1.TabIndex = 0;
            this.dataExcel1.Text = "dataExcel1";
            // 
            // frmForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(722, 476);
            this.Controls.Add(this.dataExcel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.ResumeLayout(false);

        }

        #endregion

        private Feng.Excel.DataExcelControl dataExcel1;
    }
}