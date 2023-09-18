using Feng.Data.MsSQL;
using Feng.Excel.Interfaces;
using Feng.Forms;
using Feng.Forms.Controls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Feng.DataDesign
{
    public partial class frmMain2 : BaseForm
    {

        public void InitIDTool()
        {
            toolBarID.BarItemHeader.Visable = false;
            ToolBarItem toolBarItem = null;

            toolBarItem = new ToolBarItem();
            toolBarItem.Text = "刷新";
            toolBarItem.ID = "IDFresh";
            toolBarItem.Image = Feng.DataDesign.Properties.Resources.image16_gem_options;
            toolBarID.Items.AddItem(toolBarItem);
            toolBarID.ItemClick += ToolBarID_ItemClick; 
        }
  
        private void ToolBarID_ItemClick(object sender, ToolBarItem item)
        {
            try
            { 
                if (item.ID == "IDFresh")
                {
                    RefreshID();
                }
            }
            catch (System.Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain_Database", "ToolBarDatabase_ItemClick", ex);
            }
        }

        public void RefreshID()
        {

            Feng.Forms.Controls.TreeView.DataTreeNode node = dataTreeViewID.TreeView.Nodes.GetNodeByTag("ID");
            if (node == null)
            {
                node = dataTreeViewID.TreeView.Nodes.Add("ID");
                node.Tag = "ID";
            }
            List<ICell> ids = this.dataExcel1.EditView.IDCells.GetCells();
            node.Nodes.Clear();
            for (int i = 0; i < ids.Count; i++)
            {
                ICell cell = ids[i];
                Feng.Forms.Controls.TreeView.DataTreeNode treenode = node.Nodes.Add(cell.ID);
                treenode.Tag = cell;
            }
            dataTreeViewID.Visible = true;
            dataTreeViewID.TreeView.BeginRefreshNodes();
            dataTreeViewID.TreeView.EndRefreshNodes();
            dataTreeViewID.Invalidate();
            this.dataTreeViewID.Dock = DockStyle.Fill;
            this.dataTreeViewID.BackColor = Color.AliceBlue;
        }

        public void AddText(string text)
        {
            Excel.DataExcelControl dataExcel = new Excel.DataExcelControl();
            ICell cell = dataExcel.EditView.GetCellFromColumn("text", text);
        }
    }
}
