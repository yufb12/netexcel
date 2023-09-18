using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Feng.Net.Tcp;
using Feng.Excel;

namespace Feng.DataDesign
{
    public partial class frmForm : Form
    {
        public frmForm()
        {
            InitializeComponent();
        }
        public frmForm(string filename)
        {
            InitializeComponent();
            Init(filename);
        }
        public frmForm(DataExcel grid)
        {
            InitializeComponent();
            Init(grid.GetFileData());
        }
        public void Init(string filename)
        {
            this.dataExcel1.EditView.Open(filename);

            this.Text = this.dataExcel1.EditView.Title;
        }
        public void Init(byte[] data)
        {
            this.dataExcel1.EditView.Open(data);

            this.Text = this.dataExcel1.EditView.Title;
        }
        protected override void OnLoad(EventArgs e)
        {

            Feng.Excel.DataExcel dataExcel = this.dataExcel1.EditView;

            dataExcel.AllowChangedSize = false;
            dataExcel.DefaultCellFont = new System.Drawing.Font("Tahoma", 9F); 
            dataExcel.Font = new System.Drawing.Font("Tahoma", 9F);
            dataExcel.FocusBackColor = System.Drawing.Color.White;
            dataExcel.FocusForeColor = System.Drawing.SystemColors.ControlText;
            dataExcel.Location = new System.Drawing.Point(0, 0);
            dataExcel.MaxColumn = 10;
            dataExcel.MaxRow = 23; 
            dataExcel.PrintViewMode = false;
            dataExcel.ReadOnlyBackColor = System.Drawing.Color.White;
            dataExcel.ScrollStep = ((short)(3));
            dataExcel.SelectBorderWidth = 2;
            dataExcel.ShowSelectAddRect = false; 
            dataExcel.Text = "dataExcel1";

            if (dataExcel.DisplayArea != null)
            {
                int minrow = dataExcel.DisplayArea.MinCell.Row.Index;
                int maxrow = dataExcel.DisplayArea.MaxCell.Row.Index;
                int mincolumn = dataExcel.DisplayArea.MinCell.Column.Index;
                int maxcolumn = dataExcel.DisplayArea.MaxCell.Column.Index;
                float height = 0;
                for (int i = minrow; i <= maxrow; i++)
                {
                    height = height + dataExcel.Rows[i].Height;
                }
                float width = 0;
                for (int i = mincolumn; i <= maxcolumn; i++)
                {
                    width = width + dataExcel.Columns[i].Width;
                }
                dataExcel.FirstDisplayedColumnIndex = mincolumn;
                dataExcel.FirstDisplayedRowIndex = minrow;
                dataExcel.AutoShowScroller = false;
                //dataExcel.ShowColumnHeader = false;
                //dataExcel.ShowGridColumnLine = false;
                //dataExcel.ShowGridRowLine = false;
                dataExcel.ShowHorizontalRuler = false;
                dataExcel.ShowHorizontalScroller = false;
                dataExcel.ShowRowHeader = false;
                dataExcel.ShowSelectAddRect = false;
                dataExcel.ShowVerticalRuler = false;
                dataExcel.ShowVerticalScroller = false;
                if (width >= dataExcel.DefaultRowHeight)
                {
                    this.Width = (int)width;
                }
                int CaptionHeight=SystemInformation.CaptionHeight;
                if (this.FormBorderStyle == System.Windows.Forms.FormBorderStyle.None)
                {
                    CaptionHeight = 0;
                }
                this.Height = (int)height + CaptionHeight;
        
                //this.MaximumSize = new Size(this.Width, this.Height);
                //this.MinimumSize = new Size(this.Width, this.Height);
            }

            base.OnLoad(e);
        }
  

    }
}
