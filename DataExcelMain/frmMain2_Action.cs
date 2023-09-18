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
        public void InitToolBarAction()
        {
            this.toolBarCode.BarItemHeader.Visable = true; 
            Feng.Forms.Controls.ToolBarItem item = new Feng.Forms.Controls.ToolBarItem("清除", "Clear", Feng.DataDesign.Properties.Resources.CustomActionsMenu, true, true);
            this.toolBarAction.Items.Add(item); 
            this.toolBarAction.ItemClick += ToolBarAction_ItemClick;
            this.toolBarMainTool.ItemClick += ToolBarMainTool_ItemClick2;
            this.toolBarMainMenu.ItemClick+= ToolBarMainTool_ItemClick2; 
            this.txtAction.Text = string.Empty;
        }

        private void ToolBarMainTool_ItemClick2(object sender, Forms.Controls.ToolBarItem item)
        {
            try
            {
                if (item.Items.Count > 0)
                    return;
                 
                this.txtAction.AppendText(item.Text + ":"+item.ID+"\r\n");
                this.txtAction.Invalidate();
            }
            catch (Exception ex)
            {
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "ToolBarMainTool_ItemClick2", ex);
            }
        }

        private void ToolBarAction_ItemClick(object sender, Feng.Forms.Controls.ToolBarItem item)
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
                    txtCodeExecError.Text = error;
                    if (selectcodestart > 0)
                    {
                        txtCode_tabPageEvent.SelectionStart = selectcodestart;
                    }
                } 
                else if (item.ID == "FAVID")
                {
                    this.txtCode_tabPageEvent.Text = item.ToolTip;

                }
                else if (item.ID == "OutClear")
                {
                    this.txtOut.Text = string.Empty;

                }
                else if (item.ID == "Clear")
                {
                    this.txtAction.Text = string.Empty;
                } 
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
            }
        }
         
    }
}
