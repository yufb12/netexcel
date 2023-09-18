using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Feng.Excel
{
    public partial class frmBrower : Form
    {
        public frmBrower()
        {
            InitializeComponent();
        }
        protected override void OnLoad(EventArgs e)
        {
            Feng.Excel.DataExcel dataExcel = this.dataExcel1.EditView;
            dataExcel.AllowChangedSize = false; 
            dataExcel.FilterInfo = null;
            dataExcel.Font = new System.Drawing.Font("Tahoma", 9F);
            dataExcel.LineColor = System.Drawing.Color.LightSkyBlue;
            dataExcel.Location = new System.Drawing.Point(0, 24); 
            dataExcel.Password = "";
            dataExcel.RowAutoSize = false;
            dataExcel.RowBackColor = System.Drawing.SystemColors.Window;
            dataExcel.ScrollStep = ((short)(3));
            dataExcel.SelectBorderColor = System.Drawing.Color.BlueViolet;
            dataExcel.SelectChangedBorder = false;
            dataExcel.ShowSelectBorder = false;  
            dataExcel.Text = "dataExcel1";
            base.OnLoad(e);
        }
    }
}
