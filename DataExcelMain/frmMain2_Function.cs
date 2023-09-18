using Feng.Excel.Actions;
using Feng.Excel.Builder;
using Feng.Excel.Functions;
using Feng.Forms;
using Feng.Forms.Base;
using Feng.Forms.Interface;
using Feng.Script.CBEexpress;
using System;
using System.Collections;
using System.Text;
using System.Windows.Forms;

namespace Feng.DataDesign
{
    public partial class frmMain2 : BaseForm
    {
        Feng.Forms.Controls.ToolBarItem toolbarfunctionitem = null;
        public void InitToolFunctionList()
        {
            this.toolBarFun.BarItemHeader.Visable = true; 
            Feng.Forms.Controls.ToolBarItem item =
   new Feng.Forms.Controls.ToolBarItemLabel("选择函数:", "", Feng.DataDesign.Properties.Resources.MailMergeGoToNextRecord, true, false);
            this.toolBarFun.Items.Add(item);
            item.ShowToolTip = true; 

            item = new Feng.Forms.Controls.ToolBarItemDrop("函数列表", "FunctionList", Feng.DataDesign.Properties.Resources.ToolBarItemDrop, true, true);
            this.toolBarFun.Items.Add(item);
            item.ShowToolTip = false;
            toolbarfunctionitem = item;

            item = new Feng.Forms.Controls.ToolBarItemVSplit();
            this.toolBarFun.Items.Add(item);
            item.ShowToolTip = false;

            item = new Feng.Forms.Controls.ToolBarItem("新建", "NEW", Feng.DataDesign.Properties.Resources.image16_contact_blue_add, true, true);
            this.toolBarFun.Items.Add(item);
            item.ShowToolTip = false;

            item = new Feng.Forms.Controls.ToolBarItem("删除", "DELETE", Feng.DataDesign.Properties.Resources.delete16, true, true);
            this.toolBarFun.Items.Add(item);
            item.ShowToolTip = false;

            item = new Feng.Forms.Controls.ToolBarItem("保存", "SAVE", Feng.DataDesign.Properties.Resources.CommandSave, true, true);
            this.toolBarFun.Items.Add(item);
            item.ShowToolTip = false;

            item = new Feng.Forms.Controls.ToolBarItem("运行", "RUN", Feng.DataDesign.Properties.Resources.MailMergeGoToNextRecord, true, true);
            this.toolBarFun.Items.Add(item);
            item.ShowToolTip = false;

            this.toolBarFun.ItemClick += ToolBarFun_ItemClick;
            LoadFun();
        }
        public void LoadFun()
        {
            foreach (DictionaryEntry model in this.dataExcel1.EditView.FunctionList)
            {
                Feng.Forms.Controls.ToolBarItem item = new Feng.Forms.Controls.ToolBarItem(
                  model.Key.ToString(), model.Key.ToString(), null, true, false);
                FunctionItem functionItem = new FunctionItem();
                functionItem.Key = model.Key.ToString();
                functionItem.Value = model.Value.ToString();
                item.ShowToolTip = true;
                item.Tag = functionItem; 
                toolbarfunctionitem.Items.Add(item);
            }
        }
        public class FunctionItem
        {
            public string Key { get; set; }
            public string Value { get; set; }
        }
        private void ToolBarFun_ItemClick(object sender, Feng.Forms.Controls.ToolBarItem item)
        {
            try
            {
                if (item == null)
                    return;
                if (item.ID == "RUN")
                {
                    Feng.Forms.Controls.ToolBarItem barItem = this.txtCode_tabPageFunction.Tag as Feng.Forms.Controls.ToolBarItem;
                    if (barItem != null)
                    {
                        //return;
                        FunctionItem functionItem = barItem.Tag as FunctionItem;
                        if (functionItem == null)
                            return;
                    }
                    string script = this.txtCode_tabPageFunction.Text;
                    string error = string.Empty;
                    int selectcodestart = -1;
                    OutWatch outWatch = new OutWatch();
                    try
                    {
                        error = ScriptBuilder.Debug(this.dataExcel1.EditView, this.dataExcel1.EditView.FocusedCell, script, outWatch);
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
                    txtfuneror.Text = error;
                    if (selectcodestart > 0)
                    {
                        txtCode_tabPageFunction.SelectionStart = selectcodestart;
                    }
                }
                else if (item.ID == "SAVE")
                {
                    Feng.Forms.Controls.ToolBarItem barItem = this.txtCode_tabPageFunction.Tag as Feng.Forms.Controls.ToolBarItem;
                    FunctionItem functionItem = null;
                    if (barItem != null)
                    { 
                        functionItem = barItem.Tag as FunctionItem;
                    }
                    if (functionItem != null)
                    {
                        functionItem.Value = this.txtCode_tabPageFunction.Text;
                        this.dataExcel1.EditView.FunctionList[functionItem.Key] = functionItem.Value;
                    }
                    else
                    {
                        using (Feng.Forms.Dialogs.InputTextDialog dlg = new Forms.Dialogs.InputTextDialog())
                        {
                            dlg.Text = "函数名";
                            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                functionItem = new FunctionItem();
                                functionItem.Key = dlg.txtInput.Text;
                                functionItem.Value = this.txtCode_tabPageFunction.Text;
                                this.dataExcel1.EditView.FunctionList[functionItem.Key] = functionItem.Value;
                                Feng.Forms.Controls.ToolBarItem toolitem = new Feng.Forms.Controls.ToolBarItem(
                                    functionItem.Key, functionItem.Value, null, true, false);  
                                item.ShowToolTip = false;
                                item.Tag = functionItem;
                                toolitem.Tag = functionItem;
                                toolbarfunctionitem.Items.Add(toolitem);
                            }
                        }
                    }
                }
                else if (item.ID == "NEW")
                {
                    this.txtCode_tabPageFunction.Tag = null;
                } 
                else if (item.ID == "DELETE")
                {
                    Feng.Forms.Controls.ToolBarItem barItem = this.txtCode_tabPageFunction.Tag as Feng.Forms.Controls.ToolBarItem;
                    this.txtCode_tabPageFunction.Text = string.Empty;
                    this.txtCode_tabPageFunction.Tag = null;
                    if (barItem == null)
                        return;
                    toolbarfunctionitem.Items.Remove(barItem);
                    toolbarfunctionitem.Text = string.Empty;
                }
                else  
                {
                    if (toolbarfunctionitem != item)
                    {
                        FunctionItem functionItem = item.Tag as FunctionItem;
                        if (functionItem != null)
                        {
                            this.txtCode_tabPageFunction.Text = functionItem.Value;
                            this.txtCode_tabPageFunction.Tag = item;
                            toolbarfunctionitem.Text = functionItem.Key;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
            }
        }

        public void InitData(byte[] data)
        {
            this.dataexcel.Open(data);
        }
    }
}
