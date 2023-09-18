namespace Feng.Excel
{
    partial class frmGridTableCell
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
            this.dataExcel1 = new DataExcelControl();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // dataExcel1
            // 
            this.dataExcel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));  
            this.dataExcel1.Font = new System.Drawing.Font("Tahoma", 9F); 
            this.dataExcel1.Location = new System.Drawing.Point(12, 12);
            this.dataExcel1.Name = "dataExcel1"; 
            this.dataExcel1.Size = new System.Drawing.Size(690, 421);
            this.dataExcel1.TabIndex = 0;
            this.dataExcel1.Text = "dataExcel1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 440);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // frmGridTableCell
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(714, 500);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataExcel1);
            this.Name = "frmGridTableCell";
            this.Text = "frmGridTableCell";
            this.ResumeLayout(false);

        }

        #endregion

        private Feng.Excel.DataExcelControl dataExcel1;
        private System.Windows.Forms.Button button1;
    }
}