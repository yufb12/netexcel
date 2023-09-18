using Feng.Excel.Delegates;
namespace Feng.Excel
{
    partial class frmPlus
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPlus));
            this.dataExcelControl1 = new DataExcelControl();
            this.SuspendLayout();
            // 
            // dataExcel1
            //  
            this.dataExcelControl1.Font = new System.Drawing.Font("Tahoma", 9F); 
            this.dataExcelControl1.Location = new System.Drawing.Point(12, 12); 
            this.dataExcelControl1.Name = "dataExcel1";
 
            this.dataExcelControl1.Size = new System.Drawing.Size(427, 377);
            this.dataExcelControl1.TabIndex = 8;
            this.dataExcelControl1.Text = "dataExcel1"; 
            // 
            // frmPlus
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(451, 401);
            this.Controls.Add(this.dataExcelControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(459, 435);
            this.Name = "frmPlus";
            this.Text = "插件管理";
            this.ResumeLayout(false);

        }

        #endregion

        public Feng.Excel.DataExcel dataExcel1;
        public Feng.Excel.DataExcelControl dataExcelControl1;

    }
}