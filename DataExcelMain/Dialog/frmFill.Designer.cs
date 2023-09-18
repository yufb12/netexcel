namespace Feng.Excel
{
    partial class frmFill
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmFill));
            this.txtFixText = new System.Windows.Forms.TextBox();
            this.btnok = new System.Windows.Forms.Button();
            this.radioButtonFixText = new System.Windows.Forms.RadioButton();
            this.radioButtonAddNum = new System.Windows.Forms.RadioButton();
            this.radioButtonAddTime = new System.Windows.Forms.RadioButton();
            this.txtAddNum = new System.Windows.Forms.TextBox();
            this.txtAddTime = new System.Windows.Forms.TextBox();
            this.radioButtonRandom = new System.Windows.Forms.RadioButton();
            this.txtRandom = new System.Windows.Forms.TextBox();
            this.txtAddTimeUnit = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFillRowCount = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // txtFixText
            // 
            this.txtFixText.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFixText.Location = new System.Drawing.Point(170, 23);
            this.txtFixText.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtFixText.Name = "txtFixText";
            this.txtFixText.Size = new System.Drawing.Size(409, 25);
            this.txtFixText.TabIndex = 1;
            // 
            // btnok
            // 
            this.btnok.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnok.Location = new System.Drawing.Point(170, 256);
            this.btnok.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnok.Name = "btnok";
            this.btnok.Size = new System.Drawing.Size(100, 29);
            this.btnok.TabIndex = 3;
            this.btnok.Text = "确定(&O)";
            this.btnok.UseVisualStyleBackColor = true;
            this.btnok.Click += new System.EventHandler(this.btnok_Click);
            // 
            // radioButtonFixText
            // 
            this.radioButtonFixText.AutoSize = true;
            this.radioButtonFixText.Location = new System.Drawing.Point(58, 24);
            this.radioButtonFixText.Name = "radioButtonFixText";
            this.radioButtonFixText.Size = new System.Drawing.Size(88, 19);
            this.radioButtonFixText.TabIndex = 5;
            this.radioButtonFixText.Text = "固定文本";
            this.radioButtonFixText.UseVisualStyleBackColor = true;
            // 
            // radioButtonAddNum
            // 
            this.radioButtonAddNum.AutoSize = true;
            this.radioButtonAddNum.Checked = true;
            this.radioButtonAddNum.Location = new System.Drawing.Point(58, 72);
            this.radioButtonAddNum.Name = "radioButtonAddNum";
            this.radioButtonAddNum.Size = new System.Drawing.Size(88, 19);
            this.radioButtonAddNum.TabIndex = 5;
            this.radioButtonAddNum.TabStop = true;
            this.radioButtonAddNum.Text = "递增数据";
            this.radioButtonAddNum.UseVisualStyleBackColor = true;
            // 
            // radioButtonAddTime
            // 
            this.radioButtonAddTime.AutoSize = true;
            this.radioButtonAddTime.Location = new System.Drawing.Point(58, 118);
            this.radioButtonAddTime.Name = "radioButtonAddTime";
            this.radioButtonAddTime.Size = new System.Drawing.Size(88, 19);
            this.radioButtonAddTime.TabIndex = 5;
            this.radioButtonAddTime.Text = "递增时间";
            this.radioButtonAddTime.UseVisualStyleBackColor = true;
            // 
            // txtAddNum
            // 
            this.txtAddNum.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtAddNum.Location = new System.Drawing.Point(170, 71);
            this.txtAddNum.Margin = new System.Windows.Forms.Padding(4);
            this.txtAddNum.Name = "txtAddNum";
            this.txtAddNum.Size = new System.Drawing.Size(409, 25);
            this.txtAddNum.TabIndex = 1;
            this.txtAddNum.Text = "1";
            // 
            // txtAddTime
            // 
            this.txtAddTime.Location = new System.Drawing.Point(170, 117);
            this.txtAddTime.Margin = new System.Windows.Forms.Padding(4);
            this.txtAddTime.Name = "txtAddTime";
            this.txtAddTime.Size = new System.Drawing.Size(175, 25);
            this.txtAddTime.TabIndex = 1;
            this.txtAddTime.Text = "1";
            // 
            // radioButtonRandom
            // 
            this.radioButtonRandom.AutoSize = true;
            this.radioButtonRandom.Location = new System.Drawing.Point(58, 167);
            this.radioButtonRandom.Name = "radioButtonRandom";
            this.radioButtonRandom.Size = new System.Drawing.Size(88, 19);
            this.radioButtonRandom.TabIndex = 5;
            this.radioButtonRandom.Text = "随机文本";
            this.radioButtonRandom.UseVisualStyleBackColor = true;
            // 
            // txtRandom
            // 
            this.txtRandom.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtRandom.Location = new System.Drawing.Point(170, 161);
            this.txtRandom.Margin = new System.Windows.Forms.Padding(4);
            this.txtRandom.Name = "txtRandom";
            this.txtRandom.Size = new System.Drawing.Size(409, 25);
            this.txtRandom.TabIndex = 1;
            // 
            // txtAddTimeUnit
            // 
            this.txtAddTimeUnit.FormattingEnabled = true;
            this.txtAddTimeUnit.Items.AddRange(new object[] {
            "天",
            "时",
            "分",
            "秒",
            "周",
            "月",
            "年"});
            this.txtAddTimeUnit.Location = new System.Drawing.Point(352, 117);
            this.txtAddTimeUnit.Name = "txtAddTimeUnit";
            this.txtAddTimeUnit.Size = new System.Drawing.Size(59, 23);
            this.txtAddTimeUnit.TabIndex = 7;
            this.txtAddTimeUnit.Text = "天";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(282, 215);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(22, 15);
            this.label1.TabIndex = 8;
            this.label1.Text = "行";
            // 
            // txtFillRowCount
            // 
            this.txtFillRowCount.FormattingEnabled = true;
            this.txtFillRowCount.Items.AddRange(new object[] {
            "100",
            "500",
            "1000",
            "2000",
            "5000",
            "10000",
            "20000"});
            this.txtFillRowCount.Location = new System.Drawing.Point(170, 211);
            this.txtFillRowCount.Name = "txtFillRowCount";
            this.txtFillRowCount.Size = new System.Drawing.Size(107, 23);
            this.txtFillRowCount.TabIndex = 7;
            this.txtFillRowCount.Text = "500";
            // 
            // frmFill
            // 
            this.AcceptButton = this.btnok;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(592, 298);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtFillRowCount);
            this.Controls.Add(this.txtAddTimeUnit);
            this.Controls.Add(this.radioButtonRandom);
            this.Controls.Add(this.radioButtonAddTime);
            this.Controls.Add(this.radioButtonAddNum);
            this.Controls.Add(this.radioButtonFixText);
            this.Controls.Add(this.btnok);
            this.Controls.Add(this.txtAddTime);
            this.Controls.Add(this.txtAddNum);
            this.Controls.Add(this.txtRandom);
            this.Controls.Add(this.txtFixText);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "frmFill";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnok;
        public System.Windows.Forms.TextBox txtFixText;
        public System.Windows.Forms.TextBox txtAddNum;
        public System.Windows.Forms.TextBox txtAddTime;
        public System.Windows.Forms.TextBox txtRandom;
        public System.Windows.Forms.RadioButton radioButtonFixText;
        public System.Windows.Forms.RadioButton radioButtonAddNum;
        public System.Windows.Forms.RadioButton radioButtonAddTime;
        public System.Windows.Forms.RadioButton radioButtonRandom;
        public System.Windows.Forms.ComboBox txtAddTimeUnit;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox txtFillRowCount;
    }
}