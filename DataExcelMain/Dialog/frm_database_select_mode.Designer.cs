
namespace Feng.DataDesign
{
    partial class frm_database_select_mode
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
            this.dataExcel1 = new  Excel.DataExcelControl();
            this.SuspendLayout();
            // 
            // dataExcel1
            // 
            this.dataExcel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
 
            this.dataExcel1.Dock = System.Windows.Forms.DockStyle.Fill;
 
            this.dataExcel1.Font = new System.Drawing.Font("Tahoma", 9F); 
            this.dataExcel1.Location = new System.Drawing.Point(0, 0);
 
            this.dataExcel1.Name = "dataExcel1";
 
            this.dataExcel1.Size = new System.Drawing.Size(600, 500);
            this.dataExcel1.TabIndex = 0;
            this.dataExcel1.Text = "dataExcel1"; 
            // 
            // frm_database_select_mode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 500);
            this.Controls.Add(this.dataExcel1);
            this.Name = "frm_database_select_mode";
            this.Text = "查询方式";
            this.ResumeLayout(false);

        }

        #endregion

        private Excel.DataExcelControl dataExcel1;
    }
}