using Feng.Excel.Delegates;
using Feng.Excel.Interfaces;
using System;
using System.Windows.Forms;

namespace Feng.Excel
{
    public partial class frmPlus : Form
    {
        public frmPlus()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            this.dataExcelControl1.EditView.CellClick += new CellClickEventHandler(this.dataExcel1_CellClick);
            base.OnLoad(e);
        }

        private void dataExcel1_CellClick(object sender, ICell cell)
        {
#if DEBUG2
            this.Text = string.Format("Row{0},Column{1}", cell.Row.Index, cell.Column.Index);
#endif
        }

    }
}
