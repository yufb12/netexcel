using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection; 
using Feng.Excel.Extend;
using System.Drawing.Printing;
using Feng.Excel.Args;
using Feng.Excel.Edits;
using Feng.Excel.Delegates;
using Feng.Excel.Interfaces;
using Feng.IO.File;
using Feng.Excel.Collections;
using Feng.Excel.Data;
using Feng.Excel.Print;
using Feng.Excel.App;
using Feng.Excel.Base;
using Feng.Forms.Dialogs;

namespace DataExcelMain
{
    public partial class MainExcel : Form
    {
        public MainExcel()
        {
            InitializeComponent(); 
            this.dataExcel1.AllowDrop = true;
            InitButton();
        }

        private LastFiles _lastfiles = new LastFiles();
        public LastFiles LastFiles
        {
            get
            {
                return _lastfiles;
            }
        }

        public bool AllowSaveAs
        {
            get
            {
                return this.ToolStripMenuItemsaveas.Visible;
            }
            set
            {
                this.ToolStripMenuItemsaveas.Visible = value;
            }
        }

        public void Init()
        {
            this.WindowState = FormWindowState.Maximized;
            InitFunction();
            this.splitContainerTreee.SplitterDistance = 160;
            LoadleftPanel();
        }
        public void InitButton()
        { 
            toolStripButtonEnter.Enabled = false;
        }
        public void InitFunction()
        {
            this.dataExcel1.Methods.Add(new CustomFunctions(this.dataExcel1.Methods));
        }
        public bool Design { get; set; }
        public void HideDesign()
        {
            splitContainerTreee.Panel1Collapsed = true; 
        }

        void dataExcel1_CellClick(object sender, ICell cell)
        {
            dataExcel1_FouseCellChanged(sender, cell);
        }

        void dataExcel1_ExtendCellClick(object sender, ExtendCell extencell)
        {
            this.propertyGrid1.SelectedObject = extencell;
        }

        void Main_HandleCreated(object sender, EventArgs e)
        {
            VersionCheckHanedler d = new VersionCheckHanedler(VersionCheck);
            this.Invoke(d);
        }

        public MainExcel(string filename)
        {
            InitializeComponent();

            LastFiles.Add(filename);
            LastFiles.Save(lastfile);
            this.dataExcel1.Open(filename);
            LoasLastFileMenu();

        }
        private void Read(byte[] data, string pwd)
        {
            using (Feng.Excel.IO.BinaryReader reader = new Feng.Excel.IO.BinaryReader(data))
            {
                this.dataExcel1.Open(reader, pwd);
            }
        }
        public void Read(byte[] data)
        {
            data = Feng.IO.CompressHelper.GZipDecompress(data);
            Read(data, string.Empty);
        }
        public string FileName
        {
            get
            {
                return this.dataExcel1.FileName;
            }
            set
            {
                this.dataExcel1.FileName = value;
            }
        }


        private void LoasLastFileMenu()
        {
            ToolStripLastFile.DropDownItems.Clear();
            for (int i = this.LastFiles.Count - 1; i >= 0; i--)
            {
                string file = this.LastFiles[i];
                ToolStripItem item = ToolStripLastFile.DropDownItems.Add(file);
                item.Click += new EventHandler(item_Click);
            }
        }
        void itemFunction_Click(object sender, EventArgs e)
        {

            try
            {
                ToolStripItem item = sender as ToolStripItem;
                if (item != null)
                {
                    if (this.txtfunction.Focused)
                    {
                        string text = this.txtfunction.Text;
                        text = text.Insert(this.txtfunction.SelectionStart, item.Text + "( )");
                        this.txtfunction.Text = text;
                    }
                    else
                    {
                        if (this.dataExcel1.FocusedCell != null)
                        {
                            this.dataExcel1.FocusedCell.Expression = item.Text + "( )";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }
        void item_Click(object sender, EventArgs e)
        {
            ToolStripItem item = sender as ToolStripItem;
            LastFiles.Add(item.Text);
            LastFiles.Save(lastfile);
            this.dataExcel1.Open(item.Text);
            LoasLastFileMenu();
        }
        public void SetRangeProperty(List<ICell> list)
        {
            List<ICellEditControl> listctl = new List<ICellEditControl>();
            foreach (ICell cell in list)
            {
                if (cell.OwnEditControl != null)
                {
                    listctl.Add(cell.OwnEditControl);
                }
            }
            List<object> listobj = new List<object>();
            listobj.AddRange(list.ToArray());
            object[] objs = listobj.ToArray();

            List<object> listobj2 = new List<object>();
            listobj2.AddRange(listctl.ToArray());
            object[] objs2 = listobj2.ToArray();
            lck = true;
            if (objs.Length > 0)
            {
                this.propertyGrid1.SelectedObjects = objs;
            }
            if (objs2.Length > 0)
            {
                this.propertyGrid2.SelectedObjects = objs2;
            }
            else
            {
                if (this.dataExcel1.FocusedCell != null)
                {
                    this.propertyGrid2.SelectedObject = this.dataExcel1.FocusedCell.OwnEditControl;
                }
            }
        }
        void dataExcel1_MouseUp(object sender, MouseEventArgs e)
        {
            lck = true;
            try
            {
                if (this.dataExcel1.SelectCells == null)
                    return;
                List<ICell> list = this.dataExcel1.GetSelectCells();
                SetRangeProperty(list);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
            finally
            {
                lck = false;
            }
        }

        public delegate void VersionCheckHanedler();

        public void VersionCheck()
        {
            //#if DEBUG
            //            try
            //            { 
            //                DateTime dt2 = DateTime.Parse(Product.AssemblyDateTime);
            //                if (dt > dt2)
            //                {
            //                    MessageBox.Show("最新版本为【" + version + "】,此已经不是最新版本，请下载最新版本！");
            //                    System.Diagnostics.Process.Start(Feng.Product.AssemblyHomePage);
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                MessageBox.Show("已经不是最新版本，请下载最新版本！");
            //                System.Diagnostics.Process.Start(Feng.Product.AssemblyHomePage);
            //            }
            //#endif

        }

        private void dataExcel1_CellValueChanged(object sender, CellValueChangedArgs e)
        {

        }

        #region 标准工具栏
        private void ToolStripButtonNew_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.Clear();
                this.dataExcel1.Init();
                this.dataExcel1.ReFreshFirstDisplayRowIndex();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripButtonOpen_Click(object sender, EventArgs e)
        {

            try
            {

                string file = this.dataExcel1.Open();
                if (file != string.Empty)
                {
                    this.LastFiles.Add(file);
                    this.LastFiles.Save(lastfile);
                    LoasLastFileMenu();
                }
                if (this.Design)
                {
                    this.dataExcel1.FileName = string.Empty;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripButtonSave_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.SaveEvent != null)
                {
                    this.SaveEvent(this);
                }
                else
                {
                    this.dataExcel1.Save();
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripButtonPrint_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.Print();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripButtonCut_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.CommandExcute(Feng.Excel.Commands.CommandText.CommandCut);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripButtonCopy_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.CommandExcute(Feng.Excel.Commands.CommandText.CommandCopy);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripButtonPaste_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.CommandExcute(Feng.Excel.Commands.CommandText.CommandPaste);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripButtonHelper_Click(object sender, EventArgs e)
        {

            try
            {
                System.Diagnostics.Process.Start(Product.AssemblyHomePage);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        #endregion

        #region 显示工具栏
        private void toolStripButton17_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.ShowVerticalRuler = toolStripButton17.Checked;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton18_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.ShowHorizontalRuler = toolStripButton18.Checked;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton19_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.ShowRowHeader = !toolStripButton19.Checked;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton20_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.ShowColumnHeader = !toolStripButton20.Checked;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton21_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.ShowHorizontalScroller = !toolStripButton21.Checked;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton22_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.ShowVerticalScroller = !toolStripButton22.Checked;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton23_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                splitContainer1.Panel2Collapsed = !toolStripButton23.Checked;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        #endregion

        #region 编辑工具栏
        private void toolStripButton25_Click(object sender, EventArgs e)
        {

            try
            {

                this.dataExcel1.SetSelectCellEditTextBoxCell(this.dataExcel1.GetSelectCells ());

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton26_Click(object sender, EventArgs e)
        {
            try
            {
                using (Feng.Excel.Designer.ComboxEditDesigner frm = new Feng.Excel.Designer.ComboxEditDesigner())
                {
                    frm.StartPosition = FormStartPosition.CenterScreen;
                    if (frm.ShowDialog() == DialogResult.OK)
                    {
                        this.dataExcel1.SetSelectCellEditComboBoxCell(frm.Items);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton27_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditCheckBoxCell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton28_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditRadioBoxCell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton29_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditDateTimeCell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton30_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditImageCell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton31_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditPasswordCell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton32_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditNumberCell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton44_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditCnNumberCell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }
        #endregion

        #region 设置工具栏


        private void toolStripButtonFontStyleBold_Click(object sender, EventArgs e)
        {
            try
            {
                if (toolStripButtonFontStyleBold.Checked)
                {
                    this.dataExcel1.CommandExcute("CommandFontBold");
                    //this.dataExcel1.CommandFontBold(sender, e);
                }
                else
                {
                    this.dataExcel1.CommandExcute("CommandFontBoldCancel");
                    //this.dataExcel1.CommandFontBoldCancel(sender, e);
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButtonFontStyleItalic_Click(object sender, EventArgs e)
        {
            try
            {
                if (toolStripButtonFontStyleItalic.Checked)
                {
                    this.dataExcel1.CommandExcute("CommandFontItalic");
                    //this.dataExcel1.CommandFontItalic();
                }
                else
                {
                    this.dataExcel1.CommandExcute("CommandFontItalicCancel");
                    //this.dataExcel1.CommandFontItalicCancel();
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }
        private void toolStripButtonFontStyleUnderline_Click(object sender, EventArgs e)
        {
            try
            {
                if (toolStripButtonFontStyleUnderline.Checked)
                {
                    this.dataExcel1.CommandExcute("CommandFontUnderline");
                    //this.dataExcel1.CommandFontUnderline(sender, e);
                }
                else
                {
                    this.dataExcel1.CommandExcute("CommandFontUnderlineCancel");
                    //this.dataExcel1.CommandFontUnderlineCancel(sender, e);
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButtonFontStyleStrikeout_Click(object sender, EventArgs e)
        {
            try
            {
                if (toolStripButtonFontStyleStrikeout.Checked)
                {
                    this.dataExcel1.CommandExcute("CommandFontStrikeout");
                    //this.dataExcel1.CommandFontStrikeout(sender, e);
                }
                else
                {
                    this.dataExcel1.CommandExcute("CommandFontStrikeoutCancel");
                    //this.dataExcel1.CommandFontStrikeoutCancel(sender, e);
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAlignStringLeft();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAlignStringCenter();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAlignStringRight();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton7_CheckedChanged(object sender, EventArgs e)
        {


        }

        private void toolStripButtonMergeCell_Click(object sender, EventArgs e)
        {
            try
            {
                try
                {
                    if (toolStripButtonMergeCell.Checked)
                    {
                        this.dataExcel1.CommandExcute(Feng.Excel.Commands.CommandText.CommandMergeCell);
                    }
                    else
                    {
                        this.dataExcel1.CommandExcute(Feng.Excel.Commands.CommandText.CommandMergeClear);
                    }
                }
                catch (Exception ex)
                {
                    Feng.Utils.ExceptionHelper.ShowError(ex);
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton34_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.SetSelectCellTextOrientationRotateDown(toolStripButtonDirectionVertical.Checked);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.SetSelectCellAlignLineBottom();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.SetSelectCellAlignLineCenter();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton10_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.SetSelectCellAlignLineTop();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton11_ButtonClick(object sender, EventArgs e)
        {

            try
            {
                if (this.dataExcel1.FocusedCell == null)
                    return;
                using (ColorDialog dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        if (this.dataExcel1.FocusedCell != null)
                        {
                            this.dataExcel1.SetSelectCellColorBackColor(dlg.Color);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton12_ButtonClick(object sender, EventArgs e)
        {

            try
            {
                if (this.dataExcel1.FocusedCell == null)
                    return;
                using (ColorDialog dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        if (this.dataExcel1.FocusedCell != null)
                        {
                            this.dataExcel1.SetSelectCellColorForeColor(dlg.Color);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton16_ButtonClick(object sender, EventArgs e)
        {

            try
            {
                if (this.dataExcel1.FocusedCell == null)
                    return;
                using (OpenFileDialog dlg = new OpenFileDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        this.dataExcel1.SetSelectCellImageBackImage(Bitmap.FromFile(dlg.FileName) as Bitmap);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton14_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.SetSelectCellBoarderNull();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton24_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellBorderLeftTopToRightBottom();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton15_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.CommandExcute("SetSelectCellBorderLine");
                //this.dataExcel1.SetSelectCellBorderLine();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton13_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellBorderBorderOutside();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }
        #endregion

        #region 格式工具栏

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellFormatNumberMoney();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton35_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.SetSelectCellFormatNumberPercent();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton36_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellFormatNumberThousandths();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton37_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellFormatNumberDecimalPlaces1();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton38_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellFormatNumberDecimalPlaces2();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton39_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellFormatDateTimeDay();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton40_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellFormatDateTimeg();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton41_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellFormatDateTimeG();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellFormatDateTimet();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        #endregion

        #region 表格工具栏


        private void toolStripButton43_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.dataExcel1.FocusedCell == null)
                {
                    return;
                }
                //using (frmMSSql dlg = new frmMSSql())
                //{
                //    dlg.StartPosition = FormStartPosition.CenterScreen;
                //    if (dlg.ShowDialog() == DialogResult.OK)
                //    {
                //        if (this.dataExcel1.FocusedCell.OwnDataTable == null)
                //        {
                //            int r = this.dataExcel1.FocusedCell.Row.Index + 20;
                //            int c = this.dataExcel1.FocusedCell.Column.Index + dlg.table.Columns.Count; 
                //            SelectCellCollection selc = new Feng.Excel.SelectCellCollection();
                //            selc.BeginCell = this.dataExcel1.FocusedCell;
                //            selc.EndCell = this.dataExcel1[r, c];

                //            IDataExcelTable table = this.dataExcel1.AddDataExcelTable(selc);
                //            table.DataSource = dlg.table;
                //        }
                //        else
                //        {
                //            this.dataExcel1.FocusedCell.OwnDataTable.DataSource = dlg.table;
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {

            try
            {
                MessageBox.Show("暂未提供，请联系作者！");
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {

            try
            {
                MessageBox.Show("暂未提供，请联系作者！");
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton34_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("暂未提供，请联系作者！");
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }
        #endregion

        #region 文件菜单栏

        private void ToolStripMenuItemsaveas_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.SaveAs();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripMenuItemprintview_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.PrintView();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripMenuItemexit_Click(object sender, EventArgs e)
        {

            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        #endregion

        private void ToolStripMenuItemAbout_Click(object sender, EventArgs e)
        {

            try
            {
                new AboutBox().ShowDialog();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        void txtfunction_LostFocus(object sender, EventArgs e)
        {
            try
            {
                if (this.dataExcel1.FocusedCell == null)
                    return;
                if (this.txtfunction.Text.StartsWith("="))
                {
                    this.dataExcel1.FocusedCell.Expression = this.txtfunction.Text.TrimStart('=');
                }
                else
                {
                    this.dataExcel1.FocusedCell.Expression = string.Empty;
                    this.dataExcel1.FocusedCell.Value = this.txtfunction.Text;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }
        bool lck = false;
        private void dataExcel1_FouseCellChanged(object sender, ICell cell)
        {

            try
            {
                lck = true;
                if (cell != null)
                {

                    this.txtcell.Text = cell.Name;
                    if (!string.IsNullOrEmpty(cell.Expression))
                    {
                        this.txtfunction.Text = "=" + cell.Expression;
                    }
                    else
                    {
                        this.txtfunction.Text = cell.Text;
                    }
                    this.toolStripfont.Text = cell.Font.FontFamily.Name;
                    if (cell.Row.Index < 1)
                    {
                        this.txtCellID.Text = cell.Column.FieldName;
                        
                    }
                    else
                    {
                        this.txtCellID.Text = cell.ID;
                        this.txtCellCaption.Text = cell.Caption;
                    }
                    toolStripfontsize.Text = cell.Font.Size.ToString();

                    toolStripButtonMergeCell.Checked = cell.IsMergeCell;
                    if (FontStyle.Bold == (cell.Font.Style & FontStyle.Bold))
                    {
                        toolStripButtonFontStyleBold.Checked = true;
                    }
                    else
                    {
                        toolStripButtonFontStyleBold.Checked = false;
                    }

                    if (FontStyle.Italic == (cell.Font.Style & FontStyle.Italic))
                    {
                        toolStripButtonFontStyleItalic.Checked = true;
                    }
                    else
                    {
                        toolStripButtonFontStyleItalic.Checked = false;
                    }
                    if (FontStyle.Underline == (cell.Font.Style & FontStyle.Underline))
                    {
                        toolStripButtonFontStyleUnderline.Checked = true;
                    }
                    else
                    {
                        toolStripButtonFontStyleUnderline.Checked = false;
                    }
                    toolbtnCommandTextAutoMultLine.Checked = cell.AutoMultiline;
                    if (FontStyle.Strikeout == (cell.Font.Style & FontStyle.Strikeout))
                    {
                        toolStripButtonFontStyleStrikeout.Checked = true;
                    }
                    else
                    {
                        toolStripButtonFontStyleStrikeout.Checked = false;
                    }
                    this.propertyGrid1.SelectedObject = cell;
                    return;
                }
                this.txtfunction.Text = string.Empty;
                this.txtcell.Text = string.Empty;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
            finally
            {
                lck = false;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

            try
            {
                ICell cell = this.dataExcel1.GetCellByName(this.txtcell.Text);
                if (cell != null)
                {
                    this.dataExcel1.BeginReFresh();
                    this.dataExcel1.FirstDisplayedColumnIndex = (cell.Column.Index - 5);
                    this.dataExcel1.FirstDisplayedRowIndex = (cell.Row.Index - 5);
                    this.dataExcel1.FocusedCell = cell;
                    this.dataExcel1.EndReFresh();
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripfont_DropDown(object sender, EventArgs e)
        {

            try
            {
                if (toolStripfont.Items.Count > 0)
                    return;
                for (int i = 0; i < System.Drawing.FontFamily.Families.Length; i++)
                {
                    toolStripfont.Items.Add(System.Drawing.FontFamily.Families[i].Name);
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripfont_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lck)
                return;
            try
            {
                float fontsize = this.Font.Size;
                if (!float.TryParse(toolStripfontsize.Text, out fontsize))
                {
                    fontsize = this.Font.Size;
                    toolStripfontsize.Text = this.Font.Size.ToString();
                }

                FontStyle fs = FontStyle.Regular;
                if (toolStripButtonFontStyleBold.Checked)
                {
                    fs = fs | FontStyle.Bold;
                }
                if (toolStripButtonFontStyleItalic.Checked)
                {
                    fs = fs | FontStyle.Italic;
                }
                if (toolStripButtonFontStyleUnderline.Checked)
                {
                    fs = fs | FontStyle.Underline;
                }
                if (toolStripButtonFontStyleStrikeout.Checked)
                {
                    fs = fs | FontStyle.Strikeout;
                }
                Font font = new Font(this.toolStripfont.Text, fontsize, fs);
                this.dataExcel1.SetSelectCellFont(font);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripfontsize_TextChanged(object sender, EventArgs e)
        {
            if (lck)
                return;
            try
            {
                float fontsize = this.Font.Size;
                if (!float.TryParse(toolStripfontsize.Text, out fontsize))
                {
                    toolStripfontsize.Text = this.Font.Size.ToString();
                }
                FontStyle fs = FontStyle.Regular;
                if (toolStripButtonFontStyleBold.Checked)
                {
                    fs = fs | FontStyle.Bold;
                }
                if (toolStripButtonFontStyleItalic.Checked)
                {
                    fs = fs | FontStyle.Italic;
                }
                if (toolStripButtonFontStyleUnderline.Checked)
                {
                    fs = fs | FontStyle.Underline;
                }
                if (toolStripButtonFontStyleStrikeout.Checked)
                {
                    fs = fs | FontStyle.Strikeout;
                }
                Font font = new Font(this.toolStripfont.Text, fontsize, fs);
                this.dataExcel1.FocusedCell.Font = font;
                this.dataExcel1.SetSelectCellFont(font);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void propertyGrid1_SelectedGridItemChanged(object sender, SelectedGridItemChangedEventArgs e)
        {
            if (lck)
                return;
            lck = true;
            try
            {
                this.propertyGrid2.SelectedObject = this.propertyGrid1.SelectedGridItem.Value;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
            finally
            {
                lck = false;
            }
        }

        #region 背景工具栏
        private void toolStripButton45_Click(object sender, EventArgs e)
        {

            try
            {
                IBackCell cell = this.dataExcel1.SetBackCells();
                using (System.Windows.Forms.OpenFileDialog dlg = new System.Windows.Forms.OpenFileDialog())
                {
                    dlg.Filter = "(bmp,jpg,jpeg,png)|*.bmp;*.jpg;*.jpeg;*.png|*.bmp|*.bmp|*.jpg|*.jpg|*.jpeg|*.jpeg|*.png|*.png";
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        cell.BackImage = (Bitmap)Bitmap.FromFile(dlg.FileName);
                        this.dataExcel1.RePaint();
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }
        #endregion

        #region 添加工具栏
        private void toolStripButton46_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.dataExcel1.FocusedCell == null)
                    return;
                this.dataExcel1.Rows.Insert(this.dataExcel1.FocusedCell.Row.Index, new Row(this.dataExcel1, this.dataExcel1.FocusedCell.Row.Index));
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton47_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.dataExcel1.FocusedCell == null)
                    return;
                this.dataExcel1.Rows.RemoveAt(this.dataExcel1.FocusedCell.Row.Index);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton48_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.dataExcel1.FocusedCell == null)
                    return;
                this.dataExcel1.Columns.Insert(this.dataExcel1.FocusedCell.Column.Index, new Column(this.dataExcel1, this.dataExcel1.FocusedCell.Column.Index));
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton49_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.dataExcel1.FocusedCell == null)
                    return;
                this.dataExcel1.Columns.Remove(this.dataExcel1.FocusedCell.Column);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        #endregion

        private void dataExcel1_Exception(object sender, Exception ex)
        {
            Feng.Utils.ExceptionHelper.ShowError(ex);
        }

        private void toolStripButton51_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.dataExcel1.FocusedCell == null)
                {
                    return;
                }
                if (this.dataExcel1.FocusedCell.OwnBackCell != null)
                {
                    this.dataExcel1.DeleteBackCell(this.dataExcel1.FocusedCell.OwnBackCell);
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButtonGridShowHide_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.ShowGridColumnLine = !toolStripButtonGridShowHide.Checked;
                this.dataExcel1.ShowGridRowLine = !toolStripButtonGridShowHide.Checked;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton13_ButtonClick_1(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellBorderBorderOutside();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void aToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAllLeftLineBorder();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void bToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAllLeftLineBorder();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void cToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAllRightLineBorder();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void dToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAllBottomLineBorder();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void fToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAllTopLineBorder();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private string lastfile { get { return Feng.IO.FileHelper.GetStartUpFile(Feng.IO.FileHelper.USERDATA, "\\LastFile.lfd"); } }
        private string pluspath
        {
            get
            {
                return Feng.IO.FileHelper.GetStartUpFile(Feng.IO.FileHelper.USERDATA, "\\Plus");
            }
        }

        public void InitEvent()
        {
            //this.dataExcel1.MouseUp += new MouseEventHandler(dataExcel1_MouseUp);
            //this.dataExcel1.CellClick += new CellClickEventHandler(dataExcel1_CellClick);
            //this.HandleCreated += new EventHandler(Main_HandleCreated);
            //this.dataExcel1.ExtendCellClick += new ExtendCellClickHandler(dataExcel1_ExtendCellClick);

            this.dataExcel1.MouseUp += new MouseEventHandler(dataExcel1_MouseUp);
            this.dataExcel1.CellClick += new CellClickEventHandler(dataExcel1_CellClick);
            this.dataExcel1.SelectRangeChanged += DataExcel1_SelectRangeChanged;
            this.HandleCreated += new EventHandler(Main_HandleCreated);
            this.dataExcel1.ExtendCellClick += new ExtendCellClickHandler(dataExcel1_ExtendCellClick);
            this.txtfunction.GotFocus += new EventHandler(txtfunction_GotFocus);
            this.txtfunction.LostFocus += new EventHandler(txtfunction_LostFocus);
        }

        private void DataExcel1_SelectRangeChanged(object sender, SelectRangeCollection range)
        {
            try
            {
                if (this.dataExcel1.SelectCells == null)
                    return;
                List<ICell> list = new List<ICell>();
                list.AddRange(range.ToArray());
                SetRangeProperty(list);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        void txtfunction_GotFocus(object sender, EventArgs e)
        {

        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            try
            {
                Init();
                InitEvent();
                this.LastFiles.Load(lastfile);
                LoasLastFileMenu();
                LoadPlus(pluspath);
                if (this.dataExcel1.Methods != null)
                {
                    foreach (IMethod method in this.dataExcel1.Methods)
                    {
                        System.Windows.Forms.ToolStripMenuItem menu = new ToolStripMenuItem(method.Name);
                        System.Windows.Forms.ToolStripMenuItem menutool = new ToolStripMenuItem(method.Name);
                        ToolStripMenuItemFunction.DropDownItems.Add(menu);
                        toolStripButtonfunction_other.DropDownItems.Add(menutool);
                        foreach (IMethodInfo methodinfor in method.MethodList)
                        {
                            ToolStripItem item = menu.DropDownItems.Add(methodinfor.Name);
                            item.ToolTipText = methodinfor.Description;
                            item.Click += new EventHandler(itemFunction_Click);
                            ToolStripItem itemtool = menutool.DropDownItems.Add(methodinfor.Name);
                            itemtool.ToolTipText = methodinfor.Description;
                            itemtool.Click += new EventHandler(itemFunction_Click);
                        }
                    }
                }
                ReadToolLayout();
                this.Text = Product.AssemblyTitle;

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
            finally
            {
                Feng.Forms.SplashForm.CloseFlashForm();
            }

        }

        #region 添加工具栏
        private void 图表ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                //using (frmMSSql dlg = new frmMSSql())
                //{
                //    dlg.StartPosition = FormStartPosition.CenterScreen;
                //    if (dlg.ShowDialog() == DialogResult.OK)
                //    {
                //        IDataExcelChart chart = this.dataExcel1.AddChartCell();
                //        chart.DataSource = dlg.table;
                //    }
                //}
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }
        #endregion

        private void 图片ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                ImageCell imagecell = this.dataExcel1.AddImageCell();
                imagecell.OpenImage();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void 文本框ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.AddTextCell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton53_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditCellColumnHeader();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton33_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton52_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAllButton();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton54_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditCellRowHeader();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void 数据ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {

                //using (frmMSSql dlg = new frmMSSql())
                //{
                //    dlg.StartPosition = FormStartPosition.CenterScreen;
                //    if (dlg.ShowDialog() == DialogResult.OK)
                //    {
                //        this.dataExcel1.DataSource = dlg.table;
                //    }
                //}

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void ToolStripMenuItemUndo_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.Undo();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripMenuItemRedo_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.Redo();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void 筛选ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 只显示当前值ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.Filter(this.dataExcel1.SelectCells); ;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void 清除筛选ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.ClearFilter();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton56_Click(object sender, EventArgs e)
        {

            try
            {
                //Feng.DAL.clsemployee dal = new Feng.DAL.clsemployee();
                //this.dataExcel1.DataSource = dal.GetModelList(string.Empty);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void propertyGrid2_SelectedGridItemChanged(object sender, SelectedGridItemChangedEventArgs e)
        {
            if (lck)
                return;
            lck = true;
            try
            {
                this.propertyGrid1.SelectedObject = this.propertyGrid2.SelectedGridItem.Value;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
            finally
            {
                lck = false;
            }

        }

        private void toolStripButton57_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.PrintView();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton58_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.dataExcel1.SelectCells != null)
                {
                    this.dataExcel1.DisplayArea = new SelectCellCollection();
                    this.dataExcel1.DisplayArea.BeginCell = this.dataExcel1.SelectCells.BeginCell;
                    this.dataExcel1.DisplayArea.EndCell = this.dataExcel1.SelectCells.EndCell;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton59_Click(object sender, EventArgs e)
        {
            try
            {
                frmForm frm = new frmForm(this.dataExcel1);
                //frm.Init();
                frm.ShowDialog();
                this.dataExcel1.ContentWidth = frm.Width;
                this.dataExcel1.ContentHeight = frm.Height;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void dataExcel1_SaveFile(object sender, string filename)
        {
            try
            {

                LastFiles.Add(filename);
                LastFiles.Save(lastfile);
                LoasLastFileMenu();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }
        private void dataExcel1_OpenFiled(object sender, string filename)
        {
            try
            {

                LastFiles.Add(filename);
                LastFiles.Save(lastfile);
                LoasLastFileMenu();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton60_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditCnNumber2Cell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void 边框ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditCnNumber2Cell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton62_Click(object sender, EventArgs e)
        {
            try
            {
                using (ColorDialog dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        this.dataExcel1.SetSelectCellBorderLeftTopToRightBottomColor(dlg.Color);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton63_Click(object sender, EventArgs e)
        {
            try
            {
                using (ColorDialog dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        this.dataExcel1.SetSelectCellBorderLineColor(dlg.Color);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton64_Click(object sender, EventArgs e)
        {
            try
            {
                using (ColorDialog dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        this.dataExcel1.SetSelectCellBorderLeftTopRightbottomColor(dlg.Color);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripSplitButton1_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                using (ColorDialog dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        this.dataExcel1.SetSelectCellAllRightLineBorderColor(dlg.Color);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try
            {
                using (ColorDialog dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        this.dataExcel1.SetSelectCellAllLeftLineBorderColor(dlg.Color);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            try
            {
                using (ColorDialog dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        this.dataExcel1.SetSelectCellAllTopLineBorderColor(dlg.Color);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            try
            {
                using (ColorDialog dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        this.dataExcel1.SetSelectCellAllBottomLineBorderColor(dlg.Color);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButton68_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAllBottomLineBorderWidth(2);

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAllRightLineBorderWidth(2);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAllTopLineBorderWidth(2);

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            try
            {

                this.dataExcel1.SetSelectCellAllLeftLineBorderWidth(2);

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            try
            {

                this.dataExcel1.SetSelectCellBorderLeftTopToRightBottomWidth(2);

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellBorderLeftTopRightbottomWidth(2);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellBorderLineWidth(2);

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton65_DropDownOpening(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            try
            {
                float[] lines = new float[] { 0, 0.333f, 0.666f, 1 };
                this.dataExcel1.SetSelectCellLineBorder(false, Color.Empty
                    , false, System.Drawing.Drawing2D.LineCap.AnchorMask
                    , false, System.Drawing.Drawing2D.LineCap.ArrowAnchor
                    , false, System.Drawing.Drawing2D.DashStyle.DashDotDot
                    , false, System.Drawing.Drawing2D.DashCap.Flat
                    , true, lines
                    , false, System.Drawing.Drawing2D.LineJoin.Bevel
                    , false, 0
                    , false, System.Drawing.Drawing2D.PenAlignment.Center
                    , true, 3
                    , true, false, false, false, false, false);
            }
            catch (Exception ex)
            {

                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            try
            {
                float[] lines = new float[] { 0, 0.333f, 0.666f, 1 };
                this.dataExcel1.SetSelectCellLineBorder(false, Color.Empty
                    , false, System.Drawing.Drawing2D.LineCap.AnchorMask
                    , false, System.Drawing.Drawing2D.LineCap.ArrowAnchor
                    , false, System.Drawing.Drawing2D.DashStyle.DashDotDot
                    , false, System.Drawing.Drawing2D.DashCap.Flat
                    , true, lines
                    , false, System.Drawing.Drawing2D.LineJoin.Bevel
                    , false, 0
                    , false, System.Drawing.Drawing2D.PenAlignment.Center
                    , true, 3
                    , false, false, false, true, false, false);
            }
            catch (Exception ex)
            {

                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            try
            {
                float[] lines = new float[] { 0, 0.333f, 0.666f, 1 };
                this.dataExcel1.SetSelectCellLineBorder(false, Color.Empty
                    , false, System.Drawing.Drawing2D.LineCap.AnchorMask
                    , false, System.Drawing.Drawing2D.LineCap.ArrowAnchor
                    , false, System.Drawing.Drawing2D.DashStyle.DashDotDot
                    , false, System.Drawing.Drawing2D.DashCap.Flat
                    , true, lines
                    , false, System.Drawing.Drawing2D.LineJoin.Bevel
                    , false, 0
                    , false, System.Drawing.Drawing2D.PenAlignment.Center
                    , true, 3
                    , false, true, false, false, false, false);
            }
            catch (Exception ex)
            {

                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            try
            {
                float[] lines = new float[] { 0, 0.333f, 0.666f, 1 };
                this.dataExcel1.SetSelectCellLineBorder(false, Color.Empty
                    , false, System.Drawing.Drawing2D.LineCap.AnchorMask
                    , false, System.Drawing.Drawing2D.LineCap.ArrowAnchor
                    , false, System.Drawing.Drawing2D.DashStyle.DashDotDot
                    , false, System.Drawing.Drawing2D.DashCap.Flat
                    , true, lines
                    , false, System.Drawing.Drawing2D.LineJoin.Bevel
                    , false, 0
                    , false, System.Drawing.Drawing2D.PenAlignment.Center
                    , true, 3
                    , false, false, true, false, false, false);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        #region 自定函数

        private List<IPlus> pluss = new List<IPlus>();
        public List<IPlus> Pluss
        {
            get
            {
                return pluss;
            }
        }
        private void LoadPlus(string path)
        {
            if (!System.IO.Directory.Exists(path))
            {
                return;
            }
            string[] files = System.IO.Directory.GetFiles(path);
            foreach (string file in files)
            {
                try
                {
                    Assembly ass = Assembly.LoadFile(file);

                    Type[] ts = ass.GetTypes();
                    foreach (Type t in ts)
                    {
                        Type objt = t.GetInterface("Feng.Excel.IPlus");
                        if (objt != null)
                        {
                            IPlus plus = ass.CreateInstance(t.FullName) as IPlus;
                            plus.Load(this);
                            Pluss.Add(plus);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Feng.IO.LogHelper.Log(ex);
                }

            }
        }

        #endregion

        private void 插件管理ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                //using (frmPlus frm = new frmPlus())
                //{
                //    frm.StartPosition = FormStartPosition.CenterScreen;
                //    frm.dataExcel1.AutoGenerateColumns = false;
                //    frm.dataExcel1.MaxColumn = 4;
                //    frm.dataExcel1.AllowAdd = false; 
                //    int i = 1;
                //    frm.dataExcel1.Columns[i].FieldName = "Name";
                //    frm.dataExcel1.Columns[i].Caption = "名称";

                //    i++;
                //    frm.dataExcel1.Columns[i].FieldName = "Title";
                //    frm.dataExcel1.Columns[i].Caption = "标题";
                //    i++;
                //    frm.dataExcel1.Columns[i].FieldName = "Description";
                //    frm.dataExcel1.Columns[i].Caption = "描述";

                //    i++;
                //    frm.dataExcel1.Columns[i].FieldName = "Company";
                //    frm.dataExcel1.Columns[i].Caption = "公司";


                //    frm.dataExcel1.DataSource = this._pluss;
                //    frm.ShowDialog();
                //}
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void toolStripButton55_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellTime(this.dataExcel1.GetSelectCells());
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripStatusLabel3_Click(object sender, EventArgs e)
        {

            try
            {
                System.Diagnostics.Process.Start("http://cdn.market.hiapk.com/data/upload/2014/04_21/22/com.feng.pencheng_225601.apk");
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripMenuItemZoom500_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.Zoom = 5;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripMenuItem15_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.Zoom = 1;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripMenuItem16_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.Zoom = 2;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripMenuItem17_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.Zoom = 3;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripMenuItem18_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.Zoom = 4;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton67_Click(object sender, EventArgs e)
        {
            try
            {

                this.dataExcel1.SetSelectCellEditLabelCell();

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton69_Click(object sender, EventArgs e)
        {
            try
            {

                this.dataExcel1.SetSelectCellEditLinkLabelCell();

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void ToolStripMenuItemAuto_Click(object sender, EventArgs e)
        {

        }

        private void 颜色索引ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                new ColorForm().ShowDialog();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        public delegate void BeforeSaveHanlder(object sender, CancelEventArgs e);
        public event BeforeSaveHanlder BeforeSave;
        public virtual void OnBeforeSave(CancelEventArgs e)
        {
            if (this.BeforeSave != null)
            {
                this.BeforeSave(this, e);
            }
        }

        private void openExcelEToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                using (OpenFileDialog dlg = new OpenFileDialog())
                {
                    dlg.Filter = "*.xls|*.xls";
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        //NopiUtitls.ExcelHelper.ReadDataExcel(this.dataExcel1, dlg.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void saveExcelGToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                using (SaveFileDialog dlg = new SaveFileDialog())
                {
                    dlg.Filter = "*.xls|*.xls";
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        //NopiUtitls.ExcelHelper.SaveToExcel(this.dataExcel1, dlg.FileName, false);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void txtWebUrl_Click(object sender, EventArgs e)
        {

            try
            {
                System.Diagnostics.Process.Start("http://www.booxin.com/");
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }
#if DataProject
        #region DataProject
        private Feng.DataProject.Model.ProjectList _project = null;
        public Feng.DataProject.Model.ProjectList Project
        {
            get
            {
                return this._project;
            }

            set
            {
                this._project = value;
            }
        }
        public void InitTreeTable()
        {
            List<Feng.DataProject.Model.DataTableList> list = Server.Client.GetDataTableList(Project.ID);

            TreeNode nodeparent = this.treeView1.Nodes.Add("本模块表");
            nodeparent.Tag = "本模块表";
            TreeNode node = null;
            foreach (Feng.DataProject.Model.DataTableList m in list)
            {
                node = nodeparent.Nodes.Add(m.Name);
                node.Tag = m;

                List<Feng.DataProject.Model.ColumnView> mainlist = Server.Client.GetTabletColumnsList(m.Name);
                foreach (Feng.DataProject.Model.ColumnView model in mainlist)
                {
                    ProjectTreeNode nodecolumn = new ProjectTreeNode(model.ColumnName);
                    node.Nodes.Add(nodecolumn);
                    nodecolumn.Tag = model;
                }
            }
            nodeparent = this.treeView1.Nodes.Add("所有表");
            nodeparent.Tag = "所有表";

            node = nodeparent.Nodes.Add("所有表");
            node.Tag = "所有表";

            TreeNode noded = this.treeView1.Nodes.Add("默认值");
            noded.Tag = "默认值";
            string[] defaults = Server.GetDefaults();
            foreach (string str in defaults)
            {
                ProjectTreeNode nodecd = new ProjectTreeNode(str);
                nodecd.Tag = "默认值";
                nodecd.NodeType = ProjectEnum_TreeNode.Default;
                noded.Nodes.Add(nodecd);
            }
            InitModuleNumID();
        }
        public virtual void InitModuleNumID()
        {
            TreeNode noded = this.treeView1.Nodes.Add("自动编号");
            noded.Tag = "自动编号";
            List<DataProject.Model.NubID> list = Server.Client.GetNubIDList();
            foreach (DataProject.Model.NubID model in list)
            {
                if (!model.ID.Equals(string.Empty))
                {
                    ProjectTreeNode nodecd = new ProjectTreeNode("#" + model.Name);
                    nodecd.Tag = "自动编号";
                    nodecd.NodeType = ProjectEnum_TreeNode.AutoID;
                    noded.Nodes.Add(nodecd);
                }
            }
        }

        public string FullProjectPath { get; set; }
        #endregion
#endif



        public int GetFieldRowIndex(string name)
        {
            int index = 1;
            foreach (ICell cell in this.dataExcel1.FieldCells)
            {
                if (cell.FieldName.StartsWith(name))
                {
                    int i = cell.FieldName.IndexOf('{');
                    if (i > 0)
                    {
                        int j = cell.FieldName.IndexOf('}', i);
                        int outindex = 0;
                        string str = cell.FieldName.Substring(i + 1, j - i - 1);
                        if (!int.TryParse(str, out outindex))
                        {
                            index = 1;
                        }
                        if (index < outindex)
                        {
                            index = outindex;
                        }
                    }
                }
            }
            return index;
        }

        private void dataExcel1_DragDrop(object sender, DragEventArgs e)
        {

            try
            {
                object data = e.Data.GetData(DataFormats.Text);
                if (data != null)
                {
                    if (this.dataExcel1.DragDropCells != null)
                    {
                        SelectCellCollection dragdropcells = this.dataExcel1.DragDropCells;
                        ICell cel = dragdropcells.BeginCell;
                        string txt = data.ToString();

                        string[] strs = txt.Split('\\');
                        if (strs.Length > 0)
                        {
                            string table = strs[0];
                            string field = strs[1];

                            if (field.StartsWith("@"))
                            {
                                cel.DefaultValue = field;
                            }
                            else if (field.StartsWith("#"))
                            {
                                cel.DefaultValue = field;
                            }
                            else
                            {
                                bool res = false;

                                if (this.dataExcel1.SelectCells != null)
                                {
                                    RectangleF rect = this.dataExcel1.SelectCells.Rect;
                                    Point pt = this.dataExcel1.PointToClient(System.Windows.Forms.Control.MousePosition);
                                    if (rect.Contains(pt))
                                    {
                                        res = true;
                                    }
                                }

                                int index = GetFieldRowIndex(txt + ":");
                                if (res)
                                {
                                    ICell cell1 = this.dataExcel1.SelectCells.MinCell;
                                    ICell cell2 = this.dataExcel1.SelectCells.MaxCell;
                                    if (this.dataExcel1.SelectCells.Count == 1)
                                    {
                                        ICell cell = this.dataExcel1[cel.Row.Index, cel.Column.Index - 1];
                                        if (cell.OwnMergeCell != null)
                                        {
                                            cell = cell.OwnMergeCell;
                                        }
                                        if (cell.Column.Index > 0)
                                        {
                                            if (string.IsNullOrWhiteSpace(cell.Text)
                                                && string.IsNullOrWhiteSpace(cell.FieldName)
                                                    && string.IsNullOrWhiteSpace(cell.Expression)
                                                 && string.IsNullOrWhiteSpace(cell.ID))
                                            {
                                                cell.Value = field;
                                                cell.Text = field + ":";
                                            }
                                        }
                                    }
                                    for (int ci = cell1.Column.Index; ci <= cell2.Column.Index; ci++)
                                    {
                                        for (int i = cell1.Row.Index; i <= cell2.Row.Index; i++)
                                        {
                                            cel = this.dataExcel1[i, ci];
                                            if (cel.OwnMergeCell != null)
                                            {
                                                if (cel.OwnMergeCell.BeginCell != cel)
                                                {
                                                    continue;
                                                }
                                            }
                                            index = index + 1;
                                            cel.BorderStyle.BottomLineStyle.Visible = true;
                                            if (System.Windows.Forms.Control.ModifierKeys == Keys.Shift)
                                            {
                                                cel.FieldName = txt + ":Row{" + index + "}";
                                            }
                                            else
                                            {
                                                cel.FieldName = txt + ":Row{" + index + "}";
                                            }
                                            cel.InhertReadOnly = false;
                                            cel.ReadOnly = false;
                                        }
                                    }
                                }
                                else
                                {
                                    ICell cell = this.dataExcel1[cel.Row.Index, cel.Column.Index - 1];
                                    if (cell.OwnMergeCell != null)
                                    {
                                        cell = cell.OwnMergeCell;
                                    }
                                    if (cell.Column.Index > 0)
                                    {
                                        if (string.IsNullOrWhiteSpace(cell.Text)
                                            && string.IsNullOrWhiteSpace(cell.FieldName)
                                                && string.IsNullOrWhiteSpace(cell.Expression)
                                             && string.IsNullOrWhiteSpace(cell.ID))
                                        {
                                            cell.Value = field;
                                            cell.Text = field + ":";
                                        }
                                    }
                                    cel.BorderStyle.BottomLineStyle.Visible = true;
                                    cel.FieldName = txt + ":Row{" + index + "}";
                                    cel.InhertReadOnly = false;
                                    cel.ReadOnly = false;
                                }
                            }
                        }

                    }
                }
                this.dataExcel1.DragDropCells = null;
                this.dataExcel1.DrawDragDropCell = false;
                this.dataExcel1.Invalidate();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void dataExcel1_DragOver(object sender, DragEventArgs e)
        {

            try
            {
                Point pt = this.dataExcel1.PointToClient(new Point(e.X, e.Y));

                if (this.dataExcel1.SelectCells != null)
                {
                    if (this.dataExcel1.SelectCells.Rect.Contains(pt))
                    {
                        if (this.dataExcel1.DragDropCells != null)
                        {
                            this.dataExcel1.DragDropCells.BeginCell = this.dataExcel1.SelectCells.BeginCell;
                            this.dataExcel1.DragDropCells.EndCell = this.dataExcel1.SelectCells.EndCell;
                            this.dataExcel1.Invalidate();
                        }
                        else
                        {
                            this.dataExcel1.DragDropCells = new SelectCellCollection();
                            this.dataExcel1.DragDropCells.BeginCell = this.dataExcel1.SelectCells.BeginCell;
                            this.dataExcel1.DragDropCells.EndCell = this.dataExcel1.SelectCells.EndCell;
                            this.dataExcel1.Invalidate();
                        }
                        return;
                    }
                }

                ICell cell = this.dataExcel1.GetCellByPoint(pt.X, pt.Y);
                if (cell != null)
                {
                    this.dataExcel1.DrawDragDropCell = true;
                    if (this.dataExcel1.DragDropCells != null)
                    {
                        if (this.dataExcel1.DragDropCells.BeginCell != cell)
                        {
                            this.dataExcel1.DragDropCells.BeginCell = cell;
                            this.dataExcel1.DragDropCells.EndCell = cell;
                            this.dataExcel1.Invalidate();
                        }
                    }
                    else
                    {
                        this.dataExcel1.DragDropCells = new SelectCellCollection();
                        this.dataExcel1.DragDropCells.BeginCell = cell;
                        this.dataExcel1.DragDropCells.EndCell = cell;
                        this.dataExcel1.Invalidate();
                    }
                }

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }


        public void LoadleftPanel()
        {
            this.leftPanelProperty.Dock = DockStyle.Fill;
            this.leftpanelTreeField.Dock = DockStyle.Fill;
            this.leftpanelTreeField.Visible = false;
        }

        public void LoadFieldCells()
        {
            treeViewField.Nodes.Clear();
            FieldDataBaseInfo fdbi = this.dataExcel1.GetFieldDataBase();
            foreach (FieldTableInfo fti in fdbi.Tables.Values)
            {
                TreeNode nodetable = treeViewField.Nodes.Add(fti.TableName);
                foreach (FieldRowInfo fri in fti.Rows.Values)
                {
                    TreeNode noderow = nodetable.Nodes.Add(fri.Index.ToString());
                    foreach (FieldCellInfo cell in fri.Cells.Values)
                    {
                        TreeNode node = noderow.Nodes.Add(cell.ColumName);
                        node.Tag = cell;
                    }
                }
            }
        }

        private void toolStripButtonleftPanelProperty_Click(object sender, EventArgs e)
        {

            try
            {
                this.leftPanelProperty.Visible = true;
                this.leftPanelProperty.Dock = DockStyle.Fill;
                foreach (Control con in this.leftPanel.Controls)
                {
                    if (con != this.leftPanelProperty)
                    {
                        con.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButtonleftpanelTreeField_Click(object sender, EventArgs e)
        {
            try
            {
                this.leftpanelTreeField.Visible = true;
                this.leftpanelTreeField.Dock = DockStyle.Fill;
                foreach (Control con in this.leftPanel.Controls)
                {
                    if (con != this.leftpanelTreeField)
                    {
                        con.Visible = false;
                    }
                }
                LoadFieldCells();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                TreeNode node = treeViewField.SelectedNode;
                if (node != null)
                {
                    ICell cell = node.Tag as ICell;
                    if (cell != null)
                    {
                        cell.FieldName = string.Empty;
                        this.treeViewField.Nodes.Remove(node);
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void treeViewField_AfterSelect(object sender, TreeViewEventArgs e)
        {

            try
            {
                TreeNode node = treeViewField.SelectedNode;
                if (node != null)
                {
                    FieldCellInfo cell = node.Tag as FieldCellInfo;
                    if (cell != null)
                    {
                        this.dataExcel1.TempSelectRect = cell.Cell;
                        this.dataExcel1.Invalidate();

                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void 可见ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.SetSelectCellVisible(true);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void 不可见ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellVisible(false);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }


        private void 设计模式时内容可见ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.SetSelectCellDesignModeVisible();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void 设置打印区域ToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            try
            {
                SelectCellCollection sel = this.dataExcel1.SelectCells as SelectCellCollection;
                if (sel != null)
                {
                    this.dataExcel1.PrintArea = sel;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void 清除打印区域ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.PrintArea = null;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }


        private void toolStripButton_DropDownBox_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.dataExcel1.SelectCells == null)
                {
                    return;
                }
                List<ICell> list = this.dataExcel1.GetSelectCells();
                bool first = false;

                foreach (ICell cell in list)
                {
                    if (!first)
                    {
                        string tablename = Feng.Excel.DataExcel.GetTableName(cell.FieldName);
                        if (string.IsNullOrEmpty(tablename))
                        {
                            continue;
                        }
                        first = true;
                        string key = string.Format("{0}", cell.Name);
                        key = "DropDownBoxSetting" + "_" + key;
                        if (cell.OwnEditControl != null)
                        {
                            CellDropDownBox editctrl = cell.OwnEditControl as CellDropDownBox;
                            if (editctrl != null)
                            {
                                key = editctrl.Key;
                            }
                        }

                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void txtCellID_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    ICell cell = this.dataExcel1.FocusedCell;
                    if (cell != null)
                    {
                        if (cell.Row.Index == 0)
                        {
                            cell.Column.FieldName = this.txtCellID.Text;
                        }
                        else
                        {
                            cell.ID = this.txtCellID.Text;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void ToolStripMenuItem_PrintSetting_Click(object sender, EventArgs e)
        {

            try
            {
                //using (System.Windows.Forms.PrintDialog dlg = new PrintDialog())
                //{
                //    if (this.dataExcel1.PrinterSettings != null)
                //    {
                //        dlg.PrinterSettings = this.dataExcel1.PrinterSettings;
                //    }
                //    if (dlg.ShowDialog() == DialogResult.OK)
                //    {
                //        this.dataExcel1.PrinterSettings = dlg.PrinterSettings;
                //    }
                //}
                using (System.Windows.Forms.PageSetupDialog dlg = new PageSetupDialog())
                {
                    dlg.PrinterSettings = this.dataExcel1.PrinterSettings;
                    dlg.Document = this.dataExcel1.PrintDocument;

                    if (this.dataExcel1.PrintDocument == null)
                    {
                        this.dataExcel1.PrintDocument = new DataPrintDocument();

                        Size size = this.dataExcel1.GetPaperSize(this.dataExcel1.PaperName);
                        this.dataExcel1.PrintDocument.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize(this.dataExcel1.PaperName, size.Width, size.Height);
                        this.dataExcel1.PrintDocument.DefaultPageSettings.Landscape = this.dataExcel1.PrintLandScope;
                        this.dataExcel1.PrintDocument.DefaultPageSettings.Margins = this.dataExcel1.PrintMargins;
                    }

                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        this.dataExcel1.PaperName = this.dataExcel1.PrintDocument.DefaultPageSettings.PaperSize.PaperName;
                        this.dataExcel1.PrintLandScope = this.dataExcel1.PrintDocument.DefaultPageSettings.Landscape;

                        this.dataExcel1.PrintMargins = PrinterUnitConvert.Convert(dlg.PageSettings.Margins, PrinterUnit.Display, PrinterUnit.TenthsOfAMillimeter);

                        //Margins ms = new Margins(100, 100, 100, 100);
                        //if (System.Globalization.RegionInfo.CurrentRegion.IsMetric)
                        //{

                        //}
                        //Margins ms2 = PrinterUnitConvert.Convert(ms, PrinterUnit.Display, PrinterUnit.TenthsOfAMillimeter);
                        //   PrinterUnitConvert.Convert
                        //dlg.PageSettings.Margins, PrinterUnit.TenthsOfAMillimeter, PrinterUnit.Display); 
                    }
                }

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }


        private void MainExcel_SizeChanged(object sender, EventArgs e)
        {

            try
            {
                if (this.Width > 800)
                {
                    this.splitContainerTreee.SplitterDistance = 160;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }



        private void toolStripButton_clearcell_Click(object sender, EventArgs e)
        {

            try
            {
                List<ICell> list = this.dataExcel1.GetSelectCells();
                foreach (ICell cell in list)
                {
                    if (cell != null)
                    {
                        if (cell.Row != null)
                        {
                            cell.Row.Cells.Remove(cell);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void 数字ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                //using (Feng.Forms.DataProject.FillDataDialog dlg = new Forms.DataProject.FillDataDialog())
                //{
                //    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                //    {

                //    }
                //}
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
        }

        private void toolbtnReadOnly_Click(object sender, EventArgs e)
        {

            try
            {
                this.dataExcel1.SetSelectCellAllReadOnly(true);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolbtnReadOnlyCancel_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellAllReadOnly(false);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        public void ReadToolLayout()
        {
            try
            {
                string filePanelData =  Feng.IO.FileHelper.GetStartUpFile(Feng.IO.FileHelper.USERDATA, "\\ToolStripPanel.dat");
                byte[] data = null;
                if (System.IO.File.Exists(filePanelData))
                {
                    data = System.IO.File.ReadAllBytes(filePanelData);
                }
                else
                {
                    data = Feng.Office.Excel.Properties.Resources.ToolStripPanel;
                }
                using (Feng.IO.BufferReader reader = new Feng.IO.BufferReader(data))
                {
                    int length = reader.ReadInt();
                    for (int i = 0; i < length; i++)
                    {
                        int cl = reader.ReadInt();// bw.Write(tspr.Controls.Length);
                        reader.ReadInt();//  bw.Write(tspr.Bounds.Left);
                        reader.ReadInt();//    bw.Write(tspr.Bounds.Top);
                        reader.ReadInt();//   bw.Write(tspr.Bounds.Width);
                        reader.ReadInt();//  bw.Write(tspr.Bounds.Height);
                        for (int m = 0; m < cl; m++)
                        {
                            string name = reader.ReadString();// bw.Write(c.Name);
                            int cleft = reader.ReadInt();// bw.Write(c.Left);
                            int ctop = reader.ReadInt();//  bw.Write(c.Top);
                            int cwidth = reader.ReadInt();//  bw.Write(c.Width);
                            int cheight = reader.ReadInt();//  bw.Write(c.Height);
                            foreach (Control c in this.toolStripContainer1.TopToolStripPanel.Controls)
                            {
                                if (c.Name == name)
                                {
                                    c.Left = cleft;
                                    c.Top = ctop;
                                    break;
                                }
                            }
                        }
                    }
                }
               
            }
            finally
            {
            }

        }

        private void 保存工具栏布局ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                ToolStripPanelRow[] rows = this.toolStripContainer1.TopToolStripPanel.Rows;
                Feng.IO.BufferWriter bw = new Feng.IO.BufferWriter();
                int len = 0;
                len = rows.Length;
                bw.Write(len);
                foreach (ToolStripPanelRow tspr in rows)
                {
                    bw.Write(tspr.Controls.Length);
                    bw.Write(tspr.Bounds.Left);
                    bw.Write(tspr.Bounds.Top);
                    bw.Write(tspr.Bounds.Width);
                    bw.Write(tspr.Bounds.Height);
                    foreach (Control c in tspr.Controls)
                    {
                        bw.Write(c.Name);
                        bw.Write(c.Left);
                        bw.Write(c.Top);
                        bw.Write(c.Width);
                        bw.Write(c.Height);
                    }
                }
                System.IO.File.WriteAllBytes("ToolStripPanel.dat", bw.GetData());
                bw.Close();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void 加载工具栏ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                ReadToolLayout();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void toolStripButtonEdit_SingleLine_Click(object sender, EventArgs e)
        {
            try
            {

                this.dataExcel1.SetSelectCellSingleLineTextBoxEdit(this.dataExcel1.GetSelectCells());

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }
 

        private void toolStripButtonImageEdit_Click(object sender, EventArgs e)
        {
            try
            {

                this.dataExcel1.SetSelectCellImageTextBoxEdit(this.dataExcel1 .GetSelectCells ());

            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButtonedit_color_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellEditColorCell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButtonSwitchButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellSwitchCell();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }
        public Feng.Forms.EventHandlers.ObjectEventHandler SaveEvent = null;
        public virtual void ShowEnter()
        {
            toolStripButtonEnter.Enabled = true;
        }
        private void toolStripButtonEnter_Click(object sender, EventArgs e)
        {
            try
            {
                this.DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }
        private void txtCellCaption_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    ICell cell = this.dataExcel1.FocusedCell;
                    if (cell != null)
                    {
                        cell.Caption = this.txtCellCaption.Text;
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButtonColumnWidth50_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.DefaultColumnWidth = 20;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButtonColumnWidth70_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.DefaultColumnWidth = 70;
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void dataExcel1_Click(object sender, EventArgs e)
        {
            try
            {
                this.propertyGridDataExcel.SelectedObject = sender;
            }
            catch (Exception ex)
            {
                Feng.Utils.TraceHelper.WriteTrace("DataExcel", "DataExcel", "Except", ex);
            }
        }

        private void toolStripButton72_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataExcel1.SetSelectCellMoveForm(this.dataExcel1.GetSelectCells ());
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton71_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton70_Click(object sender, EventArgs e)
        {

        }

        private void toolbtnCommandTextAutoMultLine_Click(object sender, EventArgs e)
        {
            try
            {
                if (!toolbtnCommandTextAutoMultLine.Checked)
                { 
                    this.dataExcel1.CommandExcute(Feng.Excel.Commands.CommandText.CommandTextAutoMultiline );
                    toolbtnCommandTextAutoMultLine.Checked = true;
                }
                else
                {
                    this.dataExcel1.CommandExcute(Feng.Excel.Commands.CommandText.CommandTextAutoMultilineCancel);
                    toolbtnCommandTextAutoMultLine.Checked = false;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        private void toolStripButton73_Click(object sender, EventArgs e)
        {
            this.dataExcel1.ShowRowHeader = false;
            this.dataExcel1.ShowColumnHeader = false;
            this.dataExcel1.ShowGridColumnLine = false;
            this.dataExcel1.ShowGridRowLine = false;
            this.dataExcel1.ShowHorizontalRuler = false;
            this.dataExcel1.ShowHorizontalScroller = false;
            this.dataExcel1.ShowSelectAddRect = false;
            this.dataExcel1.ShowVerticalRuler = false;
            this.dataExcel1.ShowVerticalScroller = false;
            this.dataExcel1.ShowSelectBorder = false;
        }

        private void iD列表ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            splitContainerTreee.Panel1Collapsed = !splitContainerTreee.Panel1Collapsed;
        }

        private void 输出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            splitContainerPrintOut.Panel2Collapsed = !splitContainerPrintOut.Panel2Collapsed;
        }

        private void toolStripButton17_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton18_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton20_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButtonDirectionVertical_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton23_Click(object sender, EventArgs e)
        {

        }
    }
}
