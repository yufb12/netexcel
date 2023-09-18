using Feng.Excel;
using Feng.Excel.Actions;
using Feng.Excel.App;
using Feng.Excel.Builder;
using Feng.Excel.Collections;
using Feng.Excel.Commands;
using Feng.Excel.Data;
using Feng.Excel.Delegates;
using Feng.Excel.Edits;
using Feng.Excel.Extend;
using Feng.Excel.Functions;
using Feng.Excel.Interfaces;
using Feng.Forms;
using Feng.Forms.Base;
using Feng.Forms.Command;
using Feng.Forms.Interface;
using Feng.Forms.Skins;
using Feng.IO.File;
using Feng.Script.CBEexpress;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace Feng.DataDesign
{

    public partial class frmMain2 : BaseForm, IDesignForm
    {
        public frmMain2()
        {
            InitializeComponent();
            this.dataExcel1.AllowDrop = true;
        }

        private void Dataexcel_SelectCellFinished(object sender, ISelectCellCollection selectcells)
        {
            try
            {
                Feng.Forms.WaitingForm2.BeginWaiting("正在设置单元格属性，数量较多,计算中..");
                List<ICell> cells = null;
                if (!panelMain.Panel2Collapsed)
                {
                    cells = selectcells.GetAllCells();
                    if (tabControlProperty.SelectedTab ==
                     tabPagePropertyColumn)
                    {
                        List<IColumn> columns = new List<IColumn>();
                        if (selectcells != null)
                        {
                            foreach (ICell item in cells)
                            {
                                if (columns.Contains(item.Column))
                                    continue;
                                columns.Add(item.Column);
                            }
                        }
                        propertyGridColumn.SelectedObjects = columns.ToArray();
                    }

                    if (tabControlProperty.SelectedTab ==
                     tabPagePropertyRow)
                    {
                        List<IRow> columns = new List<IRow>();
                        if (selectcells != null)
                        {
                            foreach (ICell item in cells)
                            {
                                if (columns.Contains(item.Row))
                                    continue;
                                columns.Add(item.Row);
                            }
                        }
                        propertyGridRow.SelectedObjects = columns.ToArray();
                    }
                    SetRangeProperty(cells);
                }
                if (cells == null)
                {
                    cells = selectcells.GetAllCells();
                }
                Feng.Forms.WaitingForm2.UpdateText("正在单元格合并数据..");
                Calc(cells);
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("", "", "", ex);
            }
            finally
            {
                Feng.Forms.WaitingForm2.EndWaiting();
            }
        }

        public Feng.Excel.DataExcel dataexcel { get { return this.dataExcel1.EditView; } }
        private LastFiles _lastfiles = new LastFiles();
        public LastFiles LastFiles
        {
            get
            {
                return _lastfiles;
            }
        }

        public void Init()
        {

            InitImage();
            InitCustomFunction();
            this.splitContainerTreee.SplitterDistance = 160;
            LoadleftPanel();
            InitMenuTool();
            InitToolBarCode();
            InitToolBarAction();
            InitToolFunctionList();
            InitToolBarOut();
            ShowTree(this, null);
            ShowProperty(this, null);
            ShowPrintOut(this, null);
            dataTreeViewID.TreeView.FocusedNodeChanged += TreeView_FocusedNodeChanged;
            this.dataExcel1.EditView.CellTextChanged += EditView_CellTextChanged;
            this.dataExcel1.EditView.CellValueChanged += EditView_CellValueChanged; 
            this.dataExcel1.EditView.CommandExcuted += EditView_CommandExcuted;
            this.dataExcel1.EditView.BeforeCommandExcute += EditView_BeforeCommandExcute;
        }

        private void EditView_BeforeCommandExcute(object sender, Excel.Args.BeforeCommandExcuteArgs e)
        {
            try
            {
                if (e.CommandText == CommandText.CommandSave)
                {
                    edittimes = 0;
                    return;
                }
                if (e.CommandText == CommandText.CommandSaveAs)
                {
                    edittimes = 0;
                    return;
                }
                addedittime();
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmDataProjectClient", "statusSum_Click", ex);
            }
        }

        private void EditView_CommandExcuted(object sender, Excel.Args.CommandExcutedArgs e)
        {
         
        }

        private int edittimes = 0;
        private void addedittime()
        {
            edittimes++;
        }
        private void EditView_CellValueChanged(object sender, Excel.Args.CellValueChangedArgs e)
        {
            try
            {
                addedittime();
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmDataProjectClient", "statusSum_Click", ex);
            }
        }

        private void EditView_CellTextChanged(object sender, ICell cell)
        {
            try
            {
                addedittime();
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmDataProjectClient", "statusSum_Click", ex);
            }
        }

        private void TreeView_FocusedNodeChanged(object sender, Feng.Forms.Controls.TreeView.DataTreeNode node)
        {
            try
            {
                if (node == null)
                    return;
                this.GotoCell(node.Tag as ICell);
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("", "", "", ex);
            }
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            btnMax_Click(this, e);
            base.OnMouseDoubleClick(e);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            try
            {
                this.MaximumSize = new Size(Screen.PrimaryScreen.WorkingArea.Width, Screen.PrimaryScreen.WorkingArea.Height);
                int width = this.toolBarMainMenu.GetItemSize();
                if (width > this.Width - 140)
                {
                    width = this.Width - 120;
                }
                this.toolBarMainMenu.Width = width + 2;

                width = this.toolBarMainTool.GetItemSize();
                if (width > this.Width - 10)
                {
                    width = this.Width - 10;
                }
                this.toolBarMainTool.Width = width + 2;
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("", "", "OnSizeChanged", ex);
            }
            base.OnSizeChanged(e);
        }


        public void RefreshFav()
        {
            favitem.Items.Clear();
            for (int i = 1; i < 10000; i++)
            {
                ICell cell = Fav.FavCode[i, 2];
                if (!string.IsNullOrEmpty(cell.Text))
                {
                    string name = cell.Text;
                    cell = Fav.FavCode[i, 3];
                    string code = cell.Text;
                    cell = Fav.FavCode[i, 4];
                    string time = cell.Text; ;
                    Feng.Forms.Controls.ToolBarItem itemc = new Feng.Forms.Controls.ToolBarItem(name, "FAVID", null, true, false);
                    favitem.Items.Add(itemc);
                    itemc.ShowToolTip = true;
                    itemc.ToolTip = code;
                }
                else
                {
                    break;
                }
            }
        }

        public void RefreshSample()
        {
            ScriptSampleCollection scriptSamples = new ScriptSampleCollection();
            scriptSamples.Init();

            sampleitem.Items.Clear();
            foreach (ScriptSample model in scriptSamples.Samples)
            {
                Feng.Forms.Controls.ToolBarItem itemc = new Feng.Forms.Controls.ToolBarItem(model.Name, "Sample", null, true, false);
                sampleitem.Items.Add(itemc);
                itemc.ShowToolTip = true;
                itemc.ToolTip = model.Script;
                itemc.Tag = model;
            }
        }

        private void dataexcel_FocusedCellChanged(object sender, ICell cell)
        {
            try
            {
                actionitem.Items.Clear();
                List<CellPropertyAction> list = PropertyActionTools.GetCellActions(cell);
                foreach (CellPropertyAction action in list)
                {
                    Feng.Forms.Controls.ToolBarItem item = new Feng.Forms.Controls.ToolBarItem(
                      action.Descript + "[" + action.ShortName + "]", action.ActionName, null, true, false);

                    item.ShowToolTip = true;
                    item.Tag = action;
                    item.ToolTip = action.Descript;
                    actionitem.Items.Add(item);
                }

                List<DataExcelPropertyAction> listgrid = PropertyActionTools.GetGridActions(this.dataexcel);
                foreach (DataExcelPropertyAction action in listgrid)
                {
                    Feng.Forms.Controls.ToolBarItem item = new Feng.Forms.Controls.ToolBarItem(
                      action.Descript + "[" + action.ShortName + "]", action.ActionName, null, true, false);

                    item.ShowToolTip = true;
                    item.Tag = action;
                    item.ToolTip = action.Descript;
                    actionitem.Items.Add(item);
                }
                IActionList actionList = cell.OwnEditControl as IActionList;
                if (actionList != null)
                {
                    List<PropertyAction> listeditaction = actionList.GetActiones();
                    foreach (PropertyAction action in listeditaction)
                    {
                        Feng.Forms.Controls.ToolBarItem item = new Feng.Forms.Controls.ToolBarItem(
                          action.Descript + "[" + action.ShortName + "]", action.ActionName, null, true, false);

                        item.ShowToolTip = true;
                        item.Tag = action;
                        item.ToolTip = action.Descript;
                        actionitem.Items.Add(item);
                    }
                }
                if (this.txtCode_tabPageEvent.Tag is DataExcelPropertyAction)
                {

                }
                else
                {
                    this.txtCode_tabPageEvent.Text = string.Empty;
                    this.txtCode_tabPageEvent.Tag = null;

                    foreach (CellPropertyAction item in list)
                    {
                        if (!string.IsNullOrWhiteSpace(item.Script))
                        {
                            this.txtCode_tabPageEvent.Text = item.Script;
                            this.txtCode_tabPageEvent.Tag = item;
                            actionitem.Tag = item;
                            actionitem.Text = item.Descript;
                        }
                    }
                    if (this.txtCode_tabPageEvent.Tag == null)
                    {
                        foreach (CellPropertyAction item in list)
                        {
                            if (item.ActionName == "PropertyOnClick")
                            {
                                this.txtCode_tabPageEvent.Text = item.Script;
                                this.txtCode_tabPageEvent.Tag = item;
                                actionitem.Tag = item;
                                actionitem.Text = item.Descript;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "dataexcel_FocusedCellChanged", ex);
            }
        }

        public void InitCustomFunction()
        {
            this.dataexcel.Methods.Add(new CustomFunctions());
            this.dataexcel.Methods.Add(DataProjectConfigMethodContainer.DefaultMethod);
        }
        public bool Design { get; set; }
        public void HideDesign()
        {
            splitContainerTreee.Panel1Collapsed = true;
        }
        public Control FonuceControl = null;
        private void PreprocessCommandID(string commandid)
        {
            switch (commandid)
            {
                case CommandText.CommandFontBold:
                    PreprocessCommandID_CommandFontBold();
                    break;
                case CommandText.CommandFontItalic:
                    PreprocessCommandID_CommandFontItalic();
                    break;
                case CommandText.CommandFontStrikeout:
                    PreprocessCommandID_CommandFontStrikeout();
                    break;
                case CommandText.CommandFontUnderline:
                    PreprocessCommandID_CommandFontUnderline();
                    break;
                default:
                    this.dataexcel.CommandExcute(commandid);
                    break;
            }
        }

        private void PreprocessCommandID_CommandFontUnderline()
        {
            List<ICell> list = this.dataexcel.GetSelectCells();
            int boldcount = 0;

            foreach (ICell item in list)
            {
                if (item.Font.Underline)
                {
                    boldcount++;
                }
            }
            if (boldcount == list.Count)
            {
                this.dataexcel.CommandExcute(CommandText.CommandFontUnderlineCancel);
            }
            else
            {
                this.dataexcel.CommandExcute(CommandText.CommandFontUnderline);
            }
        }
        private void PreprocessCommandID_CommandFontStrikeout()
        {
            List<ICell> list = this.dataexcel.GetSelectCells();
            int boldcount = 0;

            foreach (ICell item in list)
            {
                if (item.Font.Strikeout)
                {
                    boldcount++;
                }
            }
            if (boldcount == list.Count)
            {
                this.dataexcel.CommandExcute(CommandText.CommandFontStrikeoutCancel);
            }
            else
            {
                this.dataexcel.CommandExcute(CommandText.CommandFontStrikeout);
            }
        }
        private void PreprocessCommandID_CommandFontItalic()
        {
            List<ICell> list = this.dataexcel.GetSelectCells();
            int boldcount = 0;

            foreach (ICell item in list)
            {
                if (item.Font.Italic)
                {
                    boldcount++;
                }
            }
            if (boldcount == list.Count)
            {
                this.dataexcel.CommandExcute(CommandText.CommandFontItalicCancel);
            }
            else
            {
                this.dataexcel.CommandExcute(CommandText.CommandFontItalic);
            }
        }
        private void PreprocessCommandID_CommandFontBold()
        {
            List<ICell> list = this.dataexcel.GetSelectCells();
            int boldcount = 0;

            foreach (ICell item in list)
            {
                if (item.Font.Bold)
                {
                    boldcount++;
                }
            }
            if (boldcount == list.Count)
            {
                this.dataexcel.CommandExcute(CommandText.CommandFontBoldCancel);
            }
            else
            {
                this.dataexcel.CommandExcute(CommandText.CommandFontBold);
            }
        }
        private void AppCode(string code)
        {
            Feng.Forms.Controls.EditBox edit = null;
            if (tabControlEventFunction.SelectedTab == this.tabPageEvent)
            {
                edit = this.txtCode_tabPageEvent;
            }
            else if (tabControlEventFunction.SelectedTab == this.tabPageFunction)
            {
                edit = this.txtCode_tabPageFunction;
            }
            else
            {
                edit = null;
            }
            if (edit != null)
            {
                if (edit.SelectionStart > 0)
                {
                    edit.Text = edit.Text.Insert(edit.SelectionStart, " " + code + " ");
                }
                else
                {
                    edit.Text = edit.Text + "\r\n" + code;
                }
                edit.Focus();
            }
        }
        private void ToolBarMainMenu_ItemClick(object sender, Feng.Forms.Controls.ToolBarItem item)
        {
            if (item == null)
                return;
            CommandObject command = item.Tag as CommandObject;
            IMethodInfo funtion = item.Tag as IMethodInfo;
            ExtendCommand commandextend = item.Tag as ExtendCommand;
            if (command != null)
            {
                PreprocessCommandID(item.ID);
            }
            if (funtion != null)
            {
                string express = funtion.Name;
                if (!string.IsNullOrWhiteSpace(funtion.Eg))
                {
                    express = funtion.Eg;
                }

                if (FonuceControl == this.txtCode_tabPageEvent)
                {
                    AppCode(express);
                    item.ToolBar.HideDropDown();
                }
                else
                {
                    this.dataexcel.CommandSetExpress(string.Empty, express);
                }
            }
            if (commandextend != null)
            {
                commandextend.CommandEvent(this, item);
            }
            this.dataexcel.Invalidate();
        }

        void dataexcel_CellClick(object sender, ICell cell)
        {
            dataexcel_FouseCellChanged(sender, cell);
        }

        void dataexcel_ExtendCellClick(object sender, ExtendCell extencell)
        {
            this.propertyGrid1.SelectedObject = extencell;
        }

        void Main_HandleCreated(object sender, EventArgs e)
        {
            VersionCheckHanedler d = new VersionCheckHanedler(VersionCheck);
            this.Invoke(d);
        }

        public frmMain2(string filename)
        {
            InitializeComponent();

            LastFiles.Add(filename);
            LastFiles.Save(lastfile);
            this.dataexcel.Open(filename);
            LoasLastFileMenu();

        }
 
 
        public string FileName
        {
            get
            {
                return this.dataexcel.FileName;
            }
            set
            {
                this.dataexcel.FileName = value;
            }
        }
        public string Path { get; set; }

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
                        if (this.dataexcel.FocusedCell != null)
                        {
                            this.dataexcel.FocusedCell.Expression = item.Text + "( )";
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
            this.dataexcel.Open(item.Text);
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
                if (this.IsDisposed)
                    return;
                if (objs.Length > 10000)
                {
                    if (Feng.Utils.MsgBox.ShowQuestion("选中单元格数量较多超过:" + objs.Length + ",会使右侧属性设置时较慢，是否继续？") == DialogResult.OK)
                    {
                        this.propertyGrid1.SelectedObjects = objs;
                    }
                }
                else
                {
                    this.propertyGrid1.SelectedObjects = objs;
                }
            }
            if (objs2.Length > 0)
            {
                if (!this.IsDisposed)
                {
                    this.propertyGrid2.SelectedObjects = objs2;
                }
            }
            else
            {
                if (this.dataexcel.FocusedCell != null)
                {
                    this.propertyGrid2.SelectedObject = this.dataexcel.FocusedCell.OwnEditControl;
                }
            }
        }
        void dataexcel_MouseUp(object sender, MouseEventArgs e)
        {
            lck = true;
            try
            {
                if (this.dataexcel.SelectCells == null)
                    return;
                //List<ICell> list = this.dataexcel.GetSelectCells();
                //SetRangeProperty(list);
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

        private List<ExtendCommand> toolcommands = null;

        public List<ExtendCommand> ToolCommands
        {
            get
            {
                if (toolcommands == null)
                {
                    toolcommands = new List<ExtendCommand>();
                }
                return toolcommands;
            }
        }

        public class ExtendCommand
        {
            public string CommandText { get; set; }
            public string Description { get; set; }
            public Feng.Forms.EventHandlers.ObjectValueEventHandler CommandEvent { get; set; }
            public Bitmap Image { get; set; }
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

        void txtfunction_LostFocus(object sender, EventArgs e)
        {
            try
            {
                if (this.dataexcel.FocusedCell == null)
                    return;
                this.dataexcel.FocusedCell.Expression = this.txtfunction.Text.TrimStart('=');
                this.dataexcel.FocusedCell.ExecuteExpression();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }
        bool lck = false;
        private void dataexcel_FouseCellChanged(object sender, ICell cell)
        {

            try
            {
                if (this.Disposing)
                    return;
                if (this.txtCellID.Tag != null)
                {
                    ICell tagcell = this.txtCellID.Tag as ICell;
                    if (tagcell != null)
                    {
                        tagcell.ID = this.txtCellID.Text;
                    }
                }
                if (this.txtCellCaption.Tag != null)
                {
                    ICell tagcell = this.txtCellCaption.Tag as ICell;
                    if (tagcell != null)
                    {
                        tagcell.Caption = this.txtCellCaption.Text;
                    }
                }
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
                    if (cell.Row.Index < 1)
                    {
                        this.txtCellID.Text = cell.Column.ID;

                    }
                    else
                    {
                        this.txtCellID.Text = cell.ID;
                        this.txtCellCaption.Text = cell.Caption;
                        this.txtCellID.Tag = cell;
                        this.txtCellCaption.Tag = cell;
                    } 
                    if (this.Disposing) 
                        return; 
                    if (this.IsDisposed)
                        return;
                    this.propertyGrid1.SelectedObject = cell;
                    return;
                }
                this.txtfunction.Text = string.Empty;
                this.txtcell.Text = string.Empty;
            }
            catch (Exception ex)
            {
                Feng.Utils.TraceHelper.WriteTrace("frmMain2", "frmMain2", "dataexcel_FouseCellChanged", ex);
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
                ICell cell = this.dataexcel.GetCellByName(this.txtcell.Text);
                if (cell != null)
                {
                    this.dataexcel.BeginReFresh();
                    this.dataexcel.FirstDisplayedColumnIndex = (cell.Column.Index - 5);
                    this.dataexcel.FirstDisplayedRowIndex = (cell.Row.Index - 5);
                    this.dataexcel.FocusedCell = cell;
                    this.dataexcel.EndReFresh();
                }
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

        private void dataexcel_Exception(object sender, Exception ex)
        {
            Feng.Utils.ExceptionHelper.ShowError(ex);
        }

        private string lastfile { get { return Feng.IO.FileHelper.GetStartUpFileUSER("DataExcelMain", "\\LastFile.lfd"); } }
        private string pluspath
        {
            get
            {
                return Feng.IO.FileHelper.GetStartUpFileUSER("DataExcelMain", "\\Plus");
            }
        }

        public void InitEvent()
        {
            //this.dataexcel.MouseUp += new MouseEventHandler(dataexcel_MouseUp);
            //this.dataexcel.CellClick += new CellClickEventHandler(dataexcel_CellClick);
            //this.HandleCreated += new EventHandler(Main_HandleCreated);
            //this.dataexcel.ExtendCellClick += new ExtendCellClickHandler(dataexcel_ExtendCellClick);

            this.dataExcel1.MouseUp += new MouseEventHandler(dataexcel_MouseUp);
            this.dataexcel.CellClick += new CellClickEventHandler(dataexcel_CellClick);
            this.dataexcel.SelectRangeChanged += dataexcel_SelectRangeChanged;
            this.HandleCreated += new EventHandler(Main_HandleCreated);
            this.dataexcel.ExtendCellClick += new ExtendCellClickHandler(dataexcel_ExtendCellClick);
            this.txtfunction.GotFocus += new EventHandler(txtfunction_GotFocus);
            this.txtfunction.LostFocus += new EventHandler(txtfunction_LostFocus);
        }

        private void dataexcel_SelectRangeChanged(object sender, SelectRangeCollection range)
        {
            try
            {
                if (this.dataexcel.SelectCells == null)
                    return;
                //List<ICell> list = new List<ICell>();
                //list.AddRange(range.ToArray());
                //SetRangeProperty(list);
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        }

        void txtfunction_GotFocus(object sender, EventArgs e)
        {

        }

        Feng.Forms.Events.MouseDoubleClickProxy doubleClickProxy = new Feng.Forms.Events.MouseDoubleClickProxy();
        private void frmMain_Load(object sender, EventArgs e)
        {
            try
            {
                doubleClickProxy.Init(this.panelHeader);
                SkinStyle.InitControl(this.btnClose);
                SkinStyle.InitControl(this.AppIcon);
                SkinStyle.InitControl(this.btnMaxSize);
                SkinStyle.InitControl(this.btnMinSize);
                Init();
                InitEvent();
                this.LastFiles.Load(lastfile);
                LoasLastFileMenu();
                LoadPlus(pluspath);
                this.Text = Product.AssemblyTitle;
                OnSizeChanged(e);
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

        private void propertyGrid2_SelectedGridItemChanged(object sender, SelectedGridItemChangedEventArgs e)
        {

        }

        private void dataexcel_SaveFile(object sender, string filename)
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

        private void dataexcel_OpenFiled(object sender, string filename)
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
                    Assembly ass = Assembly.LoadFrom(file);

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

        public delegate void BeforeSaveHanlder(object sender, CancelEventArgs e);
        public event BeforeSaveHanlder BeforeSave;
        public virtual void OnBeforeSave(CancelEventArgs e)
        {
            if (this.BeforeSave != null)
            {
                this.BeforeSave(this, e);
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

        public int GetFieldRowIndex(string name)
        {
            int index = 1;
            foreach (ICell cell in this.dataexcel.FieldCells)
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

        private void dataexcel_DragDrop(object sender, DragEventArgs e)
        {

            try
            {
                object data = e.Data.GetData(DataFormats.Text);
                if (data != null)
                {
                    if (this.dataexcel.DragDropCells != null)
                    {
                        SelectCellCollection dragdropcells = this.dataexcel.DragDropCells;
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

                                if (this.dataexcel.SelectCells != null)
                                {
                                    RectangleF rect = this.dataexcel.SelectCells.Rect;
                                    Point pt = this.dataexcel.PointToClient(System.Windows.Forms.Control.MousePosition);
                                    if (rect.Contains(pt))
                                    {
                                        res = true;
                                    }
                                }

                                int index = GetFieldRowIndex(txt + ":");
                                if (res)
                                {
                                    ICell cell1 = this.dataexcel.SelectCells.MinCell;
                                    ICell cell2 = this.dataexcel.SelectCells.MaxCell;
                                    if (this.dataexcel.SelectCells != null)
                                    {
                                        ICell cell = this.dataexcel.SelectCells.BeginCell;
                                        if (cel != null)
                                        {
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
                                    }
                                    for (int ci = cell1.Column.Index; ci <= cell2.Column.Index; ci++)
                                    {
                                        for (int i = cell1.Row.Index; i <= cell2.Row.Index; i++)
                                        {
                                            cel = this.dataexcel[i, ci];
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
                                    ICell cell = this.dataexcel[cel.Row.Index, cel.Column.Index - 1];
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
                this.dataexcel.DragDropCells = null;
                this.dataexcel.DrawDragDropCell = false;
                this.dataexcel.Invalidate();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }

        }

        private void dataexcel_DragOver(object sender, DragEventArgs e)
        {

            try
            {
                Point pt = this.dataexcel.PointToClient(new Point(e.X, e.Y));

                if (this.dataexcel.SelectCells != null)
                {
                    if (this.dataexcel.SelectCells.Rect.Contains(pt))
                    {
                        if (this.dataexcel.DragDropCells != null)
                        {
                            this.dataexcel.DragDropCells.BeginCell = this.dataexcel.SelectCells.BeginCell;
                            this.dataexcel.DragDropCells.EndCell = this.dataexcel.SelectCells.EndCell;
                            this.dataexcel.Invalidate();
                        }
                        else
                        {
                            this.dataexcel.DragDropCells = new SelectCellCollection();
                            this.dataexcel.DragDropCells.BeginCell = this.dataexcel.SelectCells.BeginCell;
                            this.dataexcel.DragDropCells.EndCell = this.dataexcel.SelectCells.EndCell;
                            this.dataexcel.Invalidate();
                        }
                        return;
                    }
                }

                ICell cell = this.dataexcel.GetCellByPoint(pt.X, pt.Y);
                if (cell != null)
                {
                    this.dataexcel.DrawDragDropCell = true;
                    if (this.dataexcel.DragDropCells != null)
                    {
                        if (this.dataexcel.DragDropCells.BeginCell != cell)
                        {
                            this.dataexcel.DragDropCells.BeginCell = cell;
                            this.dataexcel.DragDropCells.EndCell = cell;
                            this.dataexcel.Invalidate();
                        }
                    }
                    else
                    {
                        this.dataexcel.DragDropCells = new SelectCellCollection();
                        this.dataexcel.DragDropCells.BeginCell = cell;
                        this.dataexcel.DragDropCells.EndCell = cell;
                        this.dataexcel.Invalidate();
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
            FieldDataBaseInfo fdbi = this.dataexcel.GetFieldDataBase();
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
                        this.dataexcel.TempSelectRect = cell.Cell;
                        this.dataexcel.Invalidate();

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
                    ICell cell = this.dataexcel.FocusedCell;
                    if (cell != null)
                    {
                        if (cell.Row.Index == 0)
                        {
                            cell.Column.ID = this.txtCellID.Text;
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

        private void MainExcel_SizeChanged(object sender, EventArgs e)
        {

            try
            {
                if (this.Width > 800)
                {
                    this.splitContainerTreee.SplitterDistance = 160;
                }
                panelHeader.Invalidate();
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

        private void txtCellCaption_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    ICell cell = this.dataexcel.FocusedCell;
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

        private void dataexcel_Click(object sender, EventArgs e)
        {
            try
            {
                this.propertyGridDataExcel.SelectedObject = this.dataExcel1.EditView;
                this.FonuceControl = this.dataExcel1;
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataExcel", "DataExcel", "Except", ex);
            }
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    if (System.IO.File.Exists(this.dataexcel.FileName))
                    {
                        System.Diagnostics.Process.Start(System.IO.Path.GetDirectoryName(this.dataexcel.FileName));
                    }
                }
                else
                {
                    Feng.Utils.UnsafeNativeMethods.MoveWindow(this.Handle);
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataExcelMain", "frmMain2", "panel1_MouseDown", ex);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {      
            try
            {
                if (this.edittimes > 0)
                {
                    if (Feng.Utils.MsgBox.ShowQuestion("文件已经更改是否仍要退出？") == DialogResult.OK)
                    {
                        this.Close();
                    }
                    return;
                } 
                this.Close();
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmDataProjectClient", "statusSum_Click", ex);
            }
        }

        private void btnMax_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void panelHeader_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {

                int left = toolBarMainMenu.Right;
                left = toolBarMainMenu.Left + toolBarMainMenu.Width;
                int width = btnMinSize.Left - left;
                int textwidth = 100;
                string filename = System.IO.Path.GetFileName(this.dataexcel.FileName);
                try
                {
                    SizeF sf = panelHeader.CreateGraphics().MeasureString(filename, panelHeader.Font);
                    textwidth = (int)sf.Width;
                }
                catch (Exception)
                {
                }
                Rectangle rect = new Rectangle(left + width / 2 - textwidth / 2, btnMinSize.Top + 10, textwidth, btnMinSize.Height);
                if (rect.Contains(e.Location))
                {
                    if (System.IO.File.Exists(this.dataexcel.FileName))
                    {
                        System.Diagnostics.Process.Start(System.IO.Path.GetDirectoryName(this.dataexcel.FileName));
                    }
                }
                else
                {
                    this.btnMax_Click(sender, e);
                }
            }
            catch (Exception ex)
            {
            }

        }

        private void dataexcel_SelectCellChanged(object sender, ISelectCellCollection selectcells)
        {
            try
            {
                ICell cell = this.dataexcel.FocusedCell;
                if (cell != null)
                {
                    string text = cell.Name + "  " + cell.ID;
                    if (this.dataexcel.SelectCells != null)
                    {
                        if (this.dataexcel.SelectCells.BeginCell != this.dataexcel.SelectCells.EndCell)
                        {
                            text = text + "  " + this.dataexcel.SelectCells.BeginCell.Name;
                            text = text + ":" + this.dataexcel.SelectCells.EndCell.Name;
                            text = text + "," + (this.dataexcel.SelectCells.MaxRow() - this.dataexcel.SelectCells.MinRow() + 1);
                            text = text + "," + (this.dataexcel.SelectCells.MaxColumn() - this.dataexcel.SelectCells.MinColumn() + 1);
                        }
                    }
                    toolStatusCell.Text = text;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("", "", "", ex);
            }
        }

        public void Calc(List<ICell> cells)
        {
            decimal d = 0;
            decimal sum = 0;
            decimal count = 0;
            decimal avg = 0;
            decimal max = 0;
            decimal min = 0;
            for (int i = 0; i < cells.Count; i++)
            {
                ICell cell = cells[i];
                if (cell == null)
                    continue;
                if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Text))
                {
                    if (System.Windows.Forms.Control.ModifierKeys != Keys.Control)
                    {
                        continue;
                    }
                }
                d = Feng.Utils.ConvertHelper.ToDecimal(cell.Value);
                if (i == 0)
                {
                    count = 1;
                    avg = d;
                    max = d;
                    min = d;
                    sum = d;
                    continue;
                }
                if (max < d)
                {
                    max = d;
                }
                if (min > d)
                {
                    min = d;
                }
                sum = sum + d;
                count = count + 1;
            }
            if (count > 0)
            {
                avg = sum / count;
                string text = string.Format("合计:{0:0.##},计数:{1:0.##},平均:{2:0.##},最大值:{3:0.##},最小值:{4:0.##}", sum, count, avg, max, min);
                statusSum.Text = text;
            }
            else
            {
                statusSum.Text = string.Empty;
            }
        }

        private void tabControlProperty_Selected(object sender, TabControlEventArgs e)
        {
            try
            {
                ISelectCellCollection selectcells = this.dataexcel.SelectCells;
                if (tabControlProperty.SelectedTab ==
                 tabPagePropertyColumn)
                {
                    List<IColumn> columns = new List<IColumn>();
                    if (selectcells != null)
                    {
                        List<ICell> cells = selectcells.GetAllCells();
                        foreach (ICell item in cells)
                        {
                            if (columns.Contains(item.Column))
                                continue;
                            columns.Add(item.Column);
                        }
                    }
                    propertyGridColumn.SelectedObjects = columns.ToArray();
                }

                if (tabControlProperty.SelectedTab ==
                 tabPagePropertyRow)
                {
                    List<IRow> columns = new List<IRow>();
                    if (selectcells != null)
                    {
                        List<ICell> cells = selectcells.GetAllCells();
                        foreach (ICell item in cells)
                        {
                            if (columns.Contains(item.Row))
                                continue;
                            columns.Add(item.Row);
                        }
                    }
                    propertyGridRow.SelectedObjects = columns.ToArray();
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("", "", "", ex);
            }
        }

        public void GotoCell(ICell cell)
        {
            if (cell == null)
                return;
            this.dataexcel.BeginReFresh();
            this.dataexcel.FirstDisplayedColumnIndex = (cell.Column.Index - 5);
            this.dataexcel.FirstDisplayedRowIndex = (cell.Row.Index - 5);
            this.dataexcel.FocusedCell = cell;
            this.dataexcel.EndReFresh();
        }

        public void Open(byte[] data)
        {
            this.dataexcel.Open(data);
        }

        public void Open(string file)
        {
            this.dataexcel.Open(file);
        }

        public byte[] GetData()
        {
            return this.dataexcel.GetFileData();
        }

        public void InitDesgin()
        {
            this.dataexcel.BeforeCommandExcute -= dataexcel_BeforeCommandExcute;
            this.dataexcel.BeforeCommandExcute += dataexcel_BeforeCommandExcute;
        }

        private void dataexcel_BeforeCommandExcute(object sender, Excel.Args.BeforeCommandExcuteArgs e)
        {
            if (e.CommandText == CommandText.CommandSave)
            {
                e.Cancel = true;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        public IDesignForm New()
        {
            return new frmMain2();
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {
            ShowPrintOut(this, toolStripStatusLabel1);
        }

        private void toolStripStatusLabel2_Click(object sender, EventArgs e)
        {
            ShowProperty(this, toolStripStatusLabel2);
        }

        private void toolStripStatusLabel3_Click(object sender, EventArgs e)
        {
            ShowTree(this, toolStripStatusLabel2);
        }

        private void statusSum_Click(object sender, EventArgs e)
        {
            try
            {
                Feng.Forms.ClipboardHelper.SetText(statusSum.Text);
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

        private void panelHeader_Paint(object sender, PaintEventArgs e)
        {
            try
            {
                Color color = Feng.Drawing.ColorHelper.Light(panelHeader.BackColor);
                Rectangle rect = new Rectangle(0, 0, panelHeader.Width, 3);
                Feng.Drawing.GraphicsHelper.FillRectangleLinearGradient(e.Graphics, color, rect, System.Drawing.Drawing2D.LinearGradientMode.Vertical);
                int left = toolBarMainMenu.Right;
                left = toolBarMainMenu.Left + toolBarMainMenu.Width;
                int width = btnMinSize.Left - left;
                string filename = System.IO.Path.GetFileName(this.dataexcel.FileName);
                Feng.Drawing.GraphicsHelper.DrawText(e.Graphics, this.panelHeader.Font, filename, Color.Black, new Rectangle(left, btnMinSize.Top + 10, width, btnMinSize.Height));
                //e.Graphics.DrawString(filename, this.panelHeader.Font, Brushes.Black, new RectangleF(left, btnMinSize.Top+10, width, btnMinSize.Height));
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "panelHeader_Paint", ex);
            }
        }

        private void btnIndesign_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataexcel.InDesign = !this.dataexcel.InDesign;
                if (this.dataexcel.InDesign)
                {
                    btnIndesign.Text = "设计模式";
                }
                else
                {
                    btnIndesign.Text = "数据模式";
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "btnIndesign_Click", ex);
            }

        }
        string lastfilename = string.Empty;
        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                string filename = System.IO.Path.GetFileName(this.dataexcel.FileName);
                if (lastfilename != filename)
                {
                    if (!string.IsNullOrEmpty(filename))
                    {
                        this.Text = filename;
                    }
                    this.panelHeader.Invalidate();
                    lastfilename = filename;
                }
                if (this.dataexcel.InDesign)
                {
                    btnIndesign.Text = "设计模式";
                }
                else
                {
                    btnIndesign.Text = "数据模式";
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "btnIndesign_Click", ex);
            }
        }

        private void btnToolCreatAdd_Click(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

        private void btnToolCreatTable_Click(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

        private void txtNote_TextChanged(object sender, EventArgs e)
        {

        }

        private void AppIcon_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("www.dataexcel.cn");
            }
            catch (Exception ex)
            {
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

        private void txtCellID_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == System.Convert.ToChar(13))
                {
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmDataProjectClient", "statusSum_Click", ex);
            }
        }

        private void txtCellCaption_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == System.Convert.ToChar(13))
                {
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmDataProjectClient", "statusSum_Click", ex);
            }
        }
    }
}
