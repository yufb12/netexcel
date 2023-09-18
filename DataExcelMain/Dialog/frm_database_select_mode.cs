using Feng.Excel.Interfaces;
using Feng.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using static Feng.DataDesign.frmMain2;

namespace Feng.DataDesign
{
    public partial class frm_database_select_mode : Form
    {
        public frm_database_select_mode()
        {
            InitializeComponent();
        }
        private int columnnameindex = 2;
        private int querymodeindex = 4;
        public void Init(List<NodeTag> columns)
        {
            Feng.Excel.DataExcel dataExcel = this.dataExcel1.EditView;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowIcon = false;
            dataExcel.CellValueChanged += DataExcel1_CellValueChanged;
            dataExcel.CellClick += DataExcel1_CellClick;
            dataExcel.Columns[columnnameindex].Width = 120; 
            dataExcel.Columns[querymodeindex].Width = 120;
            dataExcel.Columns[querymodeindex+1].Width = 170;
            dataExcel.ReadOnly = true;
            dataExcel.ShowFocusedCellBorder = true;
            dataExcel.ShowSelectBorder = false;
            Feng.Excel.Edits.CellComboBox cellComboBox = new Excel.Edits.CellComboBox(dataExcel);
            cellComboBox.Items.Add(string.Empty);
            cellComboBox.Items.Add(NodeTag.eq);
            cellComboBox.Items.Add(NodeTag.like);
            cellComboBox.Items.Add(NodeTag.Leftlike);
            cellComboBox.Items.Add(NodeTag.Rightlike);
            cellComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            int row = 2;
            dataExcel[2, columnnameindex].Value = "字段名";
            dataExcel[2, querymodeindex].Value = "查询方式";
            for (int i = 0; i < columns.Count ; i++)
            {
                NodeTag columnInfo = columns[i];
                int rowindex = i + 3;
                ICell cell = dataExcel[rowindex, columnnameindex];
                cell.Value = columnInfo.ColumnName;
                cell = dataExcel[rowindex, querymodeindex];
                cell.OwnEditControl= cellComboBox;
                cell.Tag = columnInfo;
                cell.InhertReadOnly = false;
                cell.ReadOnly = false;
                cell.BorderStyle.BottomLineStyle.Visible = true;
                cell.EditMode = Enums.EditMode.ALL;
                cell.Value = NodeTag.eq;
                row = rowindex + 2;
            }
            ICell cell2 = dataExcel[row, columnnameindex];
            cell2.Value = "保存";
            cell2.BackColor = Color.DarkSlateBlue;
            cell2.HorizontalAlignment = StringAlignment.Center;
            cell2.Font = new Font(cell2.Font, FontStyle.Bold);

            cell2.ForeColor = Color.White;
        }

        protected override void OnLoad(EventArgs e)
        {
            Feng.Excel.DataExcel dataExcel = this.dataExcel1.EditView ;
            dataExcel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            dataExcel.BackGroundMode = false;
            dataExcel.DefaultCellFont = new System.Drawing.Font("Tahoma", 9F); 
            dataExcel.FocusBackColor = System.Drawing.Color.White;
            dataExcel.FocusForeColor = System.Drawing.SystemColors.ControlText;
            dataExcel.Font = new System.Drawing.Font("Tahoma", 9F);
            dataExcel.Location = new System.Drawing.Point(0, 0);
            dataExcel.MouseCaptureView = null; 
            dataExcel.ReadOnlyBackColor = System.Drawing.Color.White;
            dataExcel.ScrollStep = ((short)(3));
            dataExcel.ShowColumnHeader = false;
            dataExcel.ShowGridColumnLine = false;
            dataExcel.ShowGridRowLine = false;
            dataExcel.ShowHorizontalScroller = false;
            dataExcel.ShowRowHeader = false;
            dataExcel.ShowVerticalScroller = false; 
            dataExcel.Text = "dataExcel1";
            dataExcel.ToolTipVisible = false;
            base.OnLoad(e);
        }

        private void DataExcel1_CellClick(object sender, ICell cell)
        {
            try
            {
                if (cell.Text == "保存")
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frm_database_select_mode", "DataExcel1_CellValueChanged", ex);
            }
        }

        private void DataExcel1_CellValueChanged(object sender, Excel.Args.CellValueChangedArgs e)
        {
            try
            {
                ICell cell = e.Cell;
                NodeTag columnInfo = cell.Tag as NodeTag;
                if (columnInfo != null)
                {
                    columnInfo.QueryMode = cell.Text;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frm_database_select_mode", "DataExcel1_CellValueChanged", ex);
            }
        }
         
    }
}
