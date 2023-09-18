using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Feng.Excel
{
    public partial class frmGridTableCell : Form
    {
        public frmGridTableCell()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {    
            try
            {
                LoadSetting();
            }
            catch (Exception ex)
            {
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }

            base.OnLoad(e);
        }
        public void LoadSetting()
        {
            this.dataExcel1.EditView.DefaultCellFont = new System.Drawing.Font("Tahoma", 9F);
            this.dataExcel1.EditView.DesignerData = null;
            this.dataExcel1.EditView.Font = new System.Drawing.Font("Tahoma", 9F);
            this.dataExcel1.EditView.FocusBackColor = System.Drawing.Color.White;
            this.dataExcel1.EditView.FocusForeColor = System.Drawing.SystemColors.ControlText;
            this.dataExcel1.EditView.Location = new System.Drawing.Point(12, 12);
            this.dataExcel1.EditView.PrintViewMode = false;
            this.dataExcel1.EditView.ReadOnlyBackColor = System.Drawing.Color.White;
            this.dataExcel1.EditView.ScrollStep = ((short)(3));
        }
        private void button1_Click(object sender, EventArgs e)
        {
            
        }
    }
}
