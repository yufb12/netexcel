using Feng.Excel.Actions;
using Feng.Excel.Builder;
using Feng.Excel.Functions;
using Feng.Forms;
using Feng.Forms.Base;
using Feng.Forms.Interface;
using Feng.Script.CBEexpress;
using System;
using System.Text;
using System.Windows.Forms;

namespace Feng.DataDesign
{
    public partial class frmMain2 : BaseForm
    {
        Feng.Forms.Controls.ToolBarItem actionitem = null;
        Feng.Forms.Controls.ToolBarItem favitem = null;
        Feng.Forms.Controls.ToolBarItem sampleitem = null;
        public void InitToolBarCode()
        {
            this.toolBarCode.BarItemHeader.Visable = true;
            this.dataexcel.FocusedCellChanged += dataexcel_FocusedCellChanged;
            this.dataexcel.SelectCellChanged += dataexcel_SelectCellChanged;
            this.dataexcel.SelectCellFinished += Dataexcel_SelectCellFinished;
            Feng.Forms.Controls.ToolBarItem item =
   new Feng.Forms.Controls.ToolBarItemLabel("选择事件:", "", Feng.DataDesign.Properties.Resources.MailMergeGoToNextRecord, true, false);
            this.toolBarCode.Items.Add(item);
            item.ShowToolTip = true;
            item.ToolTip = "更多属性可参见属性ProperyAction标签";

            item = new Feng.Forms.Controls.ToolBarItemDrop("鼠标单击", "ACTION", Feng.DataDesign.Properties.Resources.ToolBarItemDrop, true, true);
            this.toolBarCode.Items.Add(item);
            item.ShowToolTip = false;
            actionitem = item;

            item = new Feng.Forms.Controls.ToolBarItemVSplit();
            this.toolBarCode.Items.Add(item);
            item.ShowToolTip = false;

            item = new Feng.Forms.Controls.ToolBarItem("保存", "SAVE", Feng.DataDesign.Properties.Resources.CommandSave, true, true);
            this.toolBarCode.Items.Add(item);
            item.ShowToolTip = false;

            item = new Feng.Forms.Controls.ToolBarItem("运行", "RUN", Feng.DataDesign.Properties.Resources.MailMergeGoToNextRecord, true, true);
            this.toolBarCode.Items.Add(item);
            item.ShowToolTip = false;

            item = new Feng.Forms.Controls.ToolBarItem("格式化", "FORMAT", Feng.DataDesign.Properties.Resources.GroupPivotChartDataAccess, true, true);
            this.toolBarCode.Items.Add(item);
            item.ShowToolTip = false;

            item = new Feng.Forms.Controls.ToolBarItem("收藏", "FAV", Feng.DataDesign.Properties.Resources.CustomActionsMenu, true, true);
            this.toolBarCode.Items.Add(item);
            item.ShowToolTip = false;

            item = new Feng.Forms.Controls.ToolBarItem("历史", "HISTORY", Feng.DataDesign.Properties.Resources.image16_wall, true, true);
            this.toolBarCode.Items.Add(item);
            item.ShowToolTip = false;
            favitem = item;


            item = new Feng.Forms.Controls.ToolBarItem("示例", "SAMPLE", Feng.DataDesign.Properties.Resources.image16_configure16, true, true);
            this.toolBarCode.Items.Add(item);
            item.ShowToolTip = false;

            sampleitem = item;

            item = new Feng.Forms.Controls.ToolBarItem("打开测试", "OPENTEST", Feng.DataDesign.Properties.Resources.image16_configure16, true, true);
            this.toolBarCode.Items.Add(item);
            item.ShowToolTip = false;

            RefreshFav();
            RefreshSample();
            this.toolBarCode.ItemClick += ToolBarCode_ItemClick;
        }
        public void InitToolBarOut()
        {
            this.toolBarOut.BarItemHeader.Visable = true;
            Feng.Forms.Controls.ToolBarItem item = null;

            item = new Feng.Forms.Controls.ToolBarItemDrop("清除", "OutClear", Feng.DataDesign.Properties.Resources.close16, true, true);
            this.toolBarOut.Items.Add(item);
            item.ShowToolTip = false;
            item.ShowImage = false;
            this.toolBarOut.ItemClick += ToolBarCode_ItemClick;
        }
        private void ToolBarCode_ItemClick(object sender, Feng.Forms.Controls.ToolBarItem item)
        {
            try
            {
                if (item == null)
                    return;
                if (item.ID == "RUN")
                {
                    IPropertyAction action = this.txtCode_tabPageEvent.Tag as IPropertyAction;

                    string script = this.txtCode_tabPageEvent.Text;
                    string error = string.Empty;
                    int selectcodestart = -1;
                    OutWatch outWatch = new OutWatch();
                    try
                    {
                        error = ScriptBuilder.Debug(this.dataexcel, this.dataexcel.FocusedCell, script, outWatch);
                    }
                    catch (Exception ex)
                    {
                        Feng.Script.CBEexpress.CBExpressException e = ex as Feng.Script.CBEexpress.CBExpressException;
                        if (e != null)
                        {
                            if (e.Token != null)
                            {
                                selectcodestart = e.Token.Position;
                            }
                        }
                        error = ex.Message;

                    }
                    txtOut.AppendText(outWatch.GetText());
                    txtCodeExecError.Text = error;
                    if (selectcodestart > 0)
                    {
                        txtCode_tabPageEvent.SelectionStart = selectcodestart;
                    }
                }
                else if (item.ID == "SAVE")
                {
                    CellPropertyAction action = this.txtCode_tabPageEvent.Tag as CellPropertyAction;
                    if (action != null)
                    {
                        action.Script = this.txtCode_tabPageEvent.Text;
                        PropertyActionTools.UpdateCellPropertyAction(action);
                        PreprocessCommandID("CommandSave");
                        return;
                    }

                    DataExcelPropertyAction action2 = this.txtCode_tabPageEvent.Tag as DataExcelPropertyAction;
                    if (action2 != null)
                    {
                        action2.Script = this.txtCode_tabPageEvent.Text;
                        PropertyActionTools.UpdateDataExcelPropertyAction(this.dataexcel, action2);
                        PreprocessCommandID("CommandSave");
                        return;
                    }
                    IPropertyAction propertyAction = this.txtCode_tabPageEvent.Tag as IPropertyAction;
                    if (propertyAction != null)
                    {
                        propertyAction.Script = this.txtCode_tabPageEvent.Text;
                    }
                    PreprocessCommandID("CommandSave");
                }
                else if (item.ID == "FAV")
                {
                    using (Feng.Forms.Dialogs.InputTextDialog dlg = new Forms.Dialogs.InputTextDialog())
                    {
                        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            Fav.AddFav(dlg.txtInput.Text, this.txtCode_tabPageEvent.Text);
                            RefreshFav();
                        }
                    }
                }
                else if (item.ID == "FAVID")
                {
                    this.txtCode_tabPageEvent.Text = item.ToolTip;

                }
                else if (item.ID == "OPENTEST")
                {
                    Feng.Excel.Forms.frmDialog dialog = new Excel.Forms.frmDialog();
                    dialog.InitData(this.dataExcel1.EditView);
                    dialog.ShowDialog();
                }
                else if (item.ID == "OutClear")
                {
                    this.txtOut.Text = string.Empty;

                }
                else if (item.ID == "Sample")
                {
                    this.txtCode_tabPageEvent.Text = this.txtCode_tabPageEvent.Text + "\r\n" + item.ToolTip;
                }
                else if (actionitem != item)
                {
                    IPropertyAction propertyAction2 = item.Tag as IPropertyAction;
                    if (propertyAction2 != null)
                    {
                        actionitem.Tag = propertyAction2;
                        actionitem.Text = propertyAction2.Descript;
                        this.txtCode_tabPageEvent.Text = propertyAction2.Script;
                        this.txtCode_tabPageEvent.Tag = propertyAction2;

                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
            }
        }

        private void txtCode_MouseClick(object sender, MouseEventArgs e)
        {
            this.FonuceControl = this.txtCode_tabPageEvent;
        }

        public class OutWatch : IOutWatch
        {
            StringBuilder stringBuilder = new StringBuilder();
            int i = 0;
            public void Write(string txt)
            {
                stringBuilder.AppendLine(string.Format("{0} {1} {2}", (i++).ToString().PadLeft(6, '0'),
                    DateTime.Now.ToString("HH:mm:ss"), txt));
            }
            public string GetText()
            {
                return stringBuilder.ToString();
            }
        }
    }
}
