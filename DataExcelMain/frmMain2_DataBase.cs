using Feng.Data.MsSQL;
using Feng.Excel.Interfaces;
using Feng.Forms;
using Feng.Forms.Controls;
using Feng.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Feng.DataDesign
{
    public partial class frmMain2 : BaseForm
    {

        public void InitDataBaseTool()
        {
            toolBarDatabase.BarItemHeader.Visable = false;
            ToolBarItem toolBarItem = new ToolBarItem();
            toolBarItem.Text = "添加连接";
            toolBarItem.ID = "AddConnection";
            toolBarItem.Image = Feng.DataDesign.Properties.Resources.image16_group_blue_add;
            toolBarDatabase.Items.AddItem(toolBarItem);

            toolBarItem = new ToolBarItem();
            toolBarItem.Text = "刷新";
            toolBarItem.ID = "DataBaseFresh";
            toolBarItem.Image = Feng.DataDesign.Properties.Resources.image16_gem_options;
            toolBarDatabase.Items.AddItem(toolBarItem);
            toolBarDatabase.ItemClick += ToolBarDatabase_ItemClick;
            treeviewcontroldatabase.TreeView.NodeDoubleClick += TreeView_NodeDoubleClick;
        }

        public int GetEmptyRow()
        {
            for (int i = 3; i < 10000; i++)
            {
                string text = Feng.Utils.ConvertHelper.ToString(this.dataexcel[i, columnname].Value);
                if (string.IsNullOrEmpty(text))
                    return i;
            }
            return -1;
        }
        int inddex = 0;
        int columnindex = 1;
        int columntype = 2;
        int columnname = 3;
        int columnsql = 4;
        int columncode = 5;
        public void InsertCode(string type, string name, string sql, string code, params string[] args)
        {
            int row = GetEmptyRow();
            if (row < 1)
                return;
            this.dataexcel[row, columntype].Value = type;
            this.dataexcel[row, columnname].Value = name;
            this.dataexcel[row, columnsql].Value = sql;
            this.dataexcel[row, columncode].Value = code;
        }
        public void InsertCode(string type, string name, string sql, string code, List<string> list)
        {
            int row = GetEmptyRow();
            if (row < 1)
                return;
            this.dataexcel[row, columntype].Value = type;
            this.dataexcel[row, columnname].Value = name;
            this.dataexcel[row, columnsql].Value = sql;
            this.dataexcel[row, columncode].Value = code;
            for (int i = 0; i < list.Count; i++)
            {
                ICell cell = this.dataexcel[row, i + columncode + 1];
                cell.Value = list[i];
                cell.Caption = list[i];
                cell.OwnEditControl = new Feng.Excel.Edits.CellCheckBox();
            }
        }

        private void TreeView_NodeDoubleClick(object sender, Forms.Controls.TreeView.DataTreeNode node)
        {
            try
            {
                ReFfreshDataTable(node);
                ReFfreshDataColumn(node);
                if (!node.HasChild)
                {
                    if (this.dataexcel.FocusedCell == null)
                        return;
                    Forms.Controls.TreeView.DataTreeNode note = this.treeviewcontroldatabase.TreeView.FocusedNode;
                    if (note != null)
                    {
                        NodeTag nodeTagColumn = note.Tag as NodeTag;
                        if (nodeTagColumn != null)
                        {
                            this.dataexcel.FocusedCell.Value = nodeTagColumn.ColumnName;
                            this.txtfunction.Text = this.dataexcel.FocusedCell.Text;

                            this.dataexcel.FocusedCell.ID = nodeTagColumn.ColumnName;
                            this.txtCellID.Text = this.dataexcel.FocusedCell.ID;

                            ICell cell = this.dataexcel.GetLeftCell(this.dataexcel.FocusedCell);
                            if (cell != null)
                            {
                                if (string.IsNullOrWhiteSpace(cell.Text))
                                {

                                    cell.Value = nodeTagColumn.ColumnName;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "TreeView_NodeDoubleClick", ex);
            }
        }

        private void ToolBarDatabase_ItemClick(object sender, ToolBarItem item)
        {
            try
            {
                if (item.ID == "AddConnection")
                {
                    using (Feng.Forms.Dialogs.InputTextDialog dlg = new Forms.Dialogs.InputTextDialog())
                    {
                        dlg.Text = "输入服务器地址";
                        dlg.txtInput.Text = "server=.;database=**;user=sa;pwd=123456";
                        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            Feng.DataDesign.Setting.Instance.Connecton = dlg.txtInput.Text;
                            Feng.DataDesign.Setting.Instance.Save();
                        }
                    }
                }
                if (item.ID == "DataBaseFresh")
                {
                    FreshDataBase();
                }
            }
            catch (System.Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain_Database", "ToolBarDatabase_ItemClick", ex);
            }
        }
        private void toolbtnSetValue_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.dataexcel.FocusedCell == null)
                    return;
                Forms.Controls.TreeView.DataTreeNode note = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (note != null)
                {
                    NodeTag nodeTagColumn = note.Tag as NodeTag;
                    if (nodeTagColumn != null)
                    {
                        this.dataexcel.FocusedCell.Value = nodeTagColumn.ColumnName;
                        this.txtfunction.Text = this.dataexcel.FocusedCell.Text;
                    }
                }

            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        private void toolbtnSetID_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.dataexcel.FocusedCell == null)
                    return;
                Forms.Controls.TreeView.DataTreeNode note = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (note != null)
                {
                    NodeTag nodeTagColumn = note.Tag as NodeTag;
                    if (nodeTagColumn != null)
                    {
                        this.dataexcel.FocusedCell.ID = nodeTagColumn.ColumnName;
                        this.txtCellID.Text = this.dataexcel.FocusedCell.ID;
                    }
                }

            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        public void FreshDataBase()
        {
            treeviewcontroldatabase.TreeView.Nodes.Clear();
            Forms.Controls.TreeView.DataTreeNode tnode = treeviewcontroldatabase.TreeView.AddNode("服务器");
            string sql = "SELECT DB_NAME() as NAME,@@SERVERNAME as SERVERNAME";
            string connection = Feng.DataDesign.Setting.Instance.Connecton;
            if (string.IsNullOrEmpty(connection))
                return;
            SqlServerDbHelper dbHelper = new SqlServerDbHelper(connection);
            DataTable table = dbHelper.Query(sql);
            tnode.Nodes.Clear();
            foreach (DataRow item in table.Rows)
            {
                string name = Feng.Utils.ConvertHelper.ToString(item["NAME"]);
                string servername = Feng.Utils.ConvertHelper.ToString(item["SERVERNAME"]);
                Forms.Controls.TreeView.DataTreeNode node = tnode.Nodes.Add(servername + "-" + name);
                NodeTag nodeTag = new NodeTag()
                {
                    Connection = connection,
                    DataBase = name,
                    TableName = string.Empty,
                    Type = 1
                };
                node.Tag = nodeTag;
            }
            treeviewcontroldatabase.TreeView.RefreshAll();
            connection = Feng.DataDesign.Setting.Instance.Connecton1;
            if (string.IsNullOrEmpty(connection))
                return;
            dbHelper = new SqlServerDbHelper(connection);
            table = dbHelper.Query(sql);
            tnode.Nodes.Clear();
            foreach (DataRow item in table.Rows)
            {
                string name = Feng.Utils.ConvertHelper.ToString(item["NAME"]);
                Forms.Controls.TreeView.DataTreeNode node =
                    new Forms.Controls.TreeView.DataTreeNode(name);
                treeviewcontroldatabase.TreeView.Nodes.Add(node);
                NodeTag nodeTag = new NodeTag()
                {
                    Connection = connection,
                    DataBase = name,
                    TableName = string.Empty,
                    Type = 1
                };
                node.Tag = nodeTag;
            }
            connection = Feng.DataDesign.Setting.Instance.Connecton2;
            if (string.IsNullOrEmpty(connection))
                return;
            dbHelper = new SqlServerDbHelper(connection);
            table = dbHelper.Query(sql);
            tnode.Nodes.Clear();
            foreach (DataRow item in table.Rows)
            {
                string name = Feng.Utils.ConvertHelper.ToString(item["NAME"]);
                Forms.Controls.TreeView.DataTreeNode node =
                    new Forms.Controls.TreeView.DataTreeNode(name);
                treeviewcontroldatabase.TreeView.Nodes.Add(node);
                NodeTag nodeTag = new NodeTag()
                {
                    Connection = connection,
                    DataBase = name,
                    TableName = string.Empty,
                    Type = 1
                };
                node.Tag = nodeTag;
            }
        }

        public void ReFfreshDataTable(Forms.Controls.TreeView.DataTreeNode tnode)
        {
            NodeTag tnodeTag = tnode.Tag as NodeTag;
            if (tnodeTag == null)
                return;
            if (tnodeTag.Type != 1)
                return;
            string sql = "SELECT [NAME] FROM SYSOBJECTS WHERE XTYPE='U'AND [NAME]<>'DTPROPERTIES' ORDER BY [NAME]";
            SqlServerDbHelper dbHelper = new SqlServerDbHelper(tnodeTag.Connection);

            DataTable table = dbHelper.Query(sql);
            tnode.Nodes.Clear();
            foreach (DataRow item in table.Rows)
            {
                string name = Feng.Utils.ConvertHelper.ToString(item["NAME"]);
                Forms.Controls.TreeView.DataTreeNode node = tnode.Nodes.Add(name);
                NodeTag nodeTag = new NodeTag()
                {
                    Connection = tnodeTag.Connection,
                    DataBase = tnodeTag.DataBase,
                    TableName = name,
                    Type = 2
                };
                node.Tag = nodeTag;
                node.ShowCheckBox = Enums.CheckStates.Yes;
            }
            tnode.Expand();
        }

        public void ReFfreshDataColumn(Forms.Controls.TreeView.DataTreeNode tnode)
        {
            NodeTag tnodeTag = tnode.Tag as NodeTag;
            if (tnodeTag == null)
                return;
            if (tnodeTag.Type != 2)
                return;
            string sql = @"SELECT

    CLMNS.NAME AS[NAME],
    CAST(ISNULL(CIK.INDEX_COLUMN_ID, 0) AS BIT) AS[INPRIMARYKEY],
	CAST(ISNULL((SELECT TOP 1 1 FROM SYS.FOREIGN_KEY_COLUMNS AS COLFK WHERE COLFK.PARENT_COLUMN_ID = CLMNS.COLUMN_ID AND COLFK.PARENT_OBJECT_ID = CLMNS.OBJECT_ID), 0) AS BIT) AS[ISFOREIGNKEY],
	USRT.NAME AS[DATATYPE],
    ISNULL(BASET.NAME, N'') AS[SYSTEMTYPE],
	CAST(CASE WHEN BASET.NAME IN(N'NCHAR', N'NVARCHAR') AND CLMNS.MAX_LENGTH <> -1 THEN CLMNS.MAX_LENGTH / 2 ELSE CLMNS.MAX_LENGTH END AS INT) AS[LENGTH],
	CAST(CLMNS.PRECISION AS INT) AS[NUMERICPRECISION],
	CAST(CLMNS.SCALE AS INT) AS[NUMERICSCALE],
	CLMNS.IS_NULLABLE AS[NULLABLE],
    CLMNS.IS_COMPUTED AS[COMPUTED],
    CLMNS.IS_IDENTITY AS[IDENTITY],
    CAST(CLMNS.IS_SPARSE AS BIT) AS[ISSPARSE],
	CAST(CLMNS.IS_COLUMN_SET AS BIT) AS[ISCOLUMNSET]
FROM SYS.TABLES AS TBL
    INNER JOIN SYS.ALL_COLUMNS AS CLMNS ON CLMNS.OBJECT_ID = TBL.OBJECT_ID
    LEFT OUTER JOIN SYS.INDEXES AS IK ON IK.OBJECT_ID = CLMNS.OBJECT_ID AND 1 = IK.IS_PRIMARY_KEY
    LEFT OUTER JOIN SYS.INDEX_COLUMNS AS CIK ON CIK.INDEX_ID = IK.INDEX_ID AND CIK.COLUMN_ID = CLMNS.COLUMN_ID AND CIK.OBJECT_ID = CLMNS.OBJECT_ID AND 0 = CIK.IS_INCLUDED_COLUMN
    LEFT OUTER JOIN SYS.TYPES AS USRT ON USRT.USER_TYPE_ID = CLMNS.USER_TYPE_ID
    LEFT OUTER JOIN SYS.TYPES AS BASET ON(BASET.USER_TYPE_ID = CLMNS.SYSTEM_TYPE_ID AND BASET.USER_TYPE_ID = BASET.SYSTEM_TYPE_ID) OR((BASET.SYSTEM_TYPE_ID = CLMNS.SYSTEM_TYPE_ID) AND(BASET.USER_TYPE_ID = CLMNS.USER_TYPE_ID) AND(BASET.IS_USER_DEFINED = 0) AND(BASET.IS_ASSEMBLY_TYPE = 1))
    LEFT OUTER JOIN SYS.XML_SCHEMA_COLLECTIONS AS XSCCLMNS ON XSCCLMNS.XML_COLLECTION_ID = CLMNS.XML_COLLECTION_ID
    LEFT OUTER JOIN SYS.SCHEMAS AS S2CLMNS ON S2CLMNS.SCHEMA_ID = XSCCLMNS.SCHEMA_ID
WHERE(TBL.NAME = '" + tnodeTag.TableName + @"' AND SCHEMA_NAME(TBL.SCHEMA_ID) = 'DBO')
ORDER BY CLMNS.COLUMN_ID ASC ";
            SqlServerDbHelper dbHelper = new SqlServerDbHelper(tnodeTag.Connection);

            DataTable table = dbHelper.Query(sql);
            tnode.Nodes.Clear();
            foreach (DataRow item in table.Rows)
            {
                string name = Feng.Utils.ConvertHelper.ToString(item["NAME"]);
                bool inprimarykey = Feng.Utils.ConvertHelper.ToBoolean(item["INPRIMARYKEY"]);
                string datatimetype = Feng.Utils.ConvertHelper.ToString(item["DATATYPE"]);
                bool isdatatime = (datatimetype.ToLower() == "datetime");
                bool identity = Feng.Utils.ConvertHelper.ToBoolean(item["IDENTITY"]);
                bool isstring = datatimetype.ToLower().Contains("char");
                bool isint = datatimetype.ToLower().Contains("int");
                bool isdecimal = datatimetype.ToLower().Contains("decimal");
                bool isbool = datatimetype.ToLower().Contains("bit");
                string primarykeytext = inprimarykey ? "Y," : "";
                string text = name + "（" + primarykeytext + datatimetype + ")";
                Forms.Controls.TreeView.DataTreeNode node =
                    tnode.Nodes.Add(text);
                NodeTag nodeTag = new NodeTag()
                {
                    Connection = tnodeTag.Connection,
                    DataBase = tnodeTag.DataBase,
                    TableName = tnodeTag.TableName,
                    ColumnName = name,
                    Type = 3,
                    PrimaryKey = inprimarykey,
                    IsDataTime = isdatatime,
                    IsString = isstring,
                    IsInt = isint,
                    IsDecimal = isdecimal,
                    IsBool = isbool,
                    IDENTITY = identity,
                    Row = item

                };
                node.Tag = nodeTag;
                node.ShowCheckBox = Enums.CheckStates.Yes;
            }
            tnode.ExpandAll();
            tnode.TreeView.Invalidate();
        }

        public void FreshScriptFile()
        {
            Forms.Controls.TreeView.DataTreeNode node = new Forms.Controls.TreeView.DataTreeNode("服务器");
            treeviewcontroldatabase.TreeView.Nodes.Add(node);
        }

        private void tooldatabasecreatecod_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        public string CreateScriptCodeNodeTage(List<NodeTag> tabs, bool fill)
        {
            return CreateScriptCodeNodeTage(tabs, fill, false);
        }
        public string CreateScriptCodeNodeTage(List<NodeTag> tabs, bool fill, bool toprowone)
        {
            StringBuilder sbcode = new StringBuilder();
            string function = "PFGetModel";
            if (fill)
            {
                function = "PFDataFill";
                if (toprowone)
                {
                    sbcode.Append(@"var table=" + function + @"("""","""",""functionID"",");
                }
                else
                {
                    sbcode.Append(@"var table=" + function + @"("""","""",""functionID"",50,");
                }
            }
            else
            {
                sbcode.Append(@"var model=" + function + @"("""","""",""functionID"",");
            }

            foreach (NodeTag item in tabs)
            {
                if (item.IsDataTime)
                {
                    sbcode.Append("CELLDATETIME(\"" + item.ColumnName + "\"),");
                }
                else if (item.IsInt)
                {
                    sbcode.Append("CELLINT(\"" + item.ColumnName + "\"),");
                }
                else if (item.IsDecimal)
                {
                    sbcode.Append("CellDecimal(\"" + item.ColumnName + "\"),");
                }
                else if (item.IsBool)
                {
                    sbcode.Append("CELLBool(\"" + item.ColumnName + "\"),");
                }
                else
                {
                    sbcode.Append("CELLText(\"" + item.ColumnName + "\"),");
                }
            }

            sbcode.Length = sbcode.Length - 1;
            sbcode.AppendLine(");");
            if (!fill)
            {
                sbcode.Append("var res=PFExecModel(model);");
            }
            return sbcode.ToString();
        }
        public string CreateScriptCodeNodeTage_Exist(List<NodeTag> tabs)
        {
            StringBuilder sbcode = new StringBuilder();
            string function = "PFExecModel";

            function = "PFDataExist";
            sbcode.Append(@"var exist=" + function + @"("""","""",""functionID"",");
 
            foreach (NodeTag item in tabs)
            {
                if (item.IsDataTime)
                {
                    sbcode.Append("CELLDATETIME(\"" + item.ColumnName + "\"),");
                }
                else if (item.IsInt)
                {
                    sbcode.Append("CELLINT(\"" + item.ColumnName + "\"),");
                }
                else if (item.IsDecimal)
                {
                    sbcode.Append("CellDecimal(\"" + item.ColumnName + "\"),");
                }
                else if (item.IsBool)
                {
                    sbcode.Append("CELLBool(\"" + item.ColumnName + "\"),");
                }
                else
                {
                    sbcode.Append("CELLText(\"" + item.ColumnName + "\"),");
                }
            }

            sbcode.Length = sbcode.Length - 1;
            sbcode.AppendLine(");"); 
            return sbcode.ToString();
        }
        private void ToolDataBase_AddAll_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("insert into [" + nodeTag.TableName + "](");
                        List<NodeTag> tabs = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            if (nodeTagColumn.IDENTITY)
                            {
                                continue;
                            }
                            sb.Append("[" + nodeTagColumn.ColumnName + "]");
                            sb.Append(",");
                        }
                        sb.Length = sb.Length - 1;
                        sb.AppendLine(")");
                        sb.Append("values(");
                        int index = 1;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            if (nodeTagColumn.IDENTITY)
                            {
                                continue;
                            }
                            tabs.Add(nodeTagColumn);

                            sb.Append("@arg" + index + ",");
                            index++;
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(")");
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, false);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_ADD", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        private void ToolDataBase_EditAll_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //UPDATE [dbo].[table1] SET [NO] = @ARG1 WHERE [ID]=@ID
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("UPDATE [dbo].[" + nodeTag.TableName + "] SET");
                        int index = 1;
                        List<NodeTag> tabs = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            if (nodeTagColumn.PrimaryKey)
                            {
                                continue;
                            }
                            tabs.Add(nodeTagColumn);
                            sb.Append("[" + nodeTagColumn.ColumnName + "]=");
                            sb.Append("@ARG" + index + ",");
                            index++;
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(" Where 1=1 ");
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            if (!nodeTagColumn.PrimaryKey)
                            {
                                continue;
                            }
                            tabs.Add(nodeTagColumn);
                            sb.Append(" AND [" + nodeTagColumn.ColumnName + "]=");
                            sb.Append("@ARG" + index + ",");
                            index++;
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, false);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_UPDATE", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        private void ToolDataBase_Del_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //DELETE TOP(1) [dbo].[table1] WHERE [ID]=@ID
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("DELETE TOP(1) [dbo].[" + nodeTag.TableName + "] ");
                        int index = 1;
                        sb.Append(" Where 1=1 ");
                        List<NodeTag> tabs = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            if (nodeTagColumn.PrimaryKey)
                            {
                                tabs.Add(nodeTagColumn);
                                sb.Append(" AND [" + nodeTagColumn.ColumnName + "]=");
                                sb.Append("@ARG" + index + ",");
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, false);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_DELETE", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        private void ToolDataBase_QueryAll_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //SELECT TOP 100 [NO],[NAME] FROM [DBO].[TABLE1] WHERE [ID]=@ID
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("SELECT TOP (@ARG1) ");
                        int index = 2;
                        List<NodeTag> tabs = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            sb.Append("[" + nodeTagColumn.ColumnName + "],");
                            index++;
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(" FROM [DBO].[" + nodeTag.TableName + "] ");
                        sb.Append(" Where 1=1 ");
                        index = 2;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                tabs.Add(nodeTagColumn);
                                sb.Append(" AND [" + nodeTagColumn.ColumnName + "]=");
                                sb.Append("@ARG" + index + ",");
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        index = 2;
                        sb.Append(" ORDER BY ");
                        int sortcount = 0;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (sortcount > 2)
                            {
                                break;
                            }
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                if (nodeTagColumn.IsDataTime)
                                {
                                    sb.Append(" [" + nodeTagColumn.ColumnName + "] DESC,");
                                }
                                else
                                {
                                    sb.Append(" [" + nodeTagColumn.ColumnName + "] ASC,");
                                }
                                sortcount++;
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, true);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_SELECT", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        private void ToolDataBase_AddSel_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("insert into [" + nodeTag.TableName + "](");
                        List<NodeTag> tabs = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                if (nodeTagColumn.IDENTITY)
                                {
                                    continue;
                                }
                                tabs.Add(nodeTagColumn);
                                sb.Append("[" + nodeTagColumn.ColumnName + "]");
                                sb.Append(",");
                            }
                        }
                        sb.Length = sb.Length - 1;
                        sb.AppendLine(")");
                        sb.Append("values(");
                        int index = 1;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                if (nodeTagColumn.IDENTITY)
                                {
                                    continue;
                                }
                                sb.Append("@arg" + index + ",");
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(")");
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, false);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_ADD", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        private void ToolDataBase_EditSel_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //UPDATE [dbo].[table1] SET [NO] = @ARG1 WHERE [ID]=@ID
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("UPDATE [dbo].[" + nodeTag.TableName + "] SET");
                        int index = 1;
                        List<NodeTag> tabs = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                if (nodeTagColumn.PrimaryKey)
                                {
                                    continue;
                                }
                                tabs.Add(nodeTagColumn);
                                sb.Append("[" + nodeTagColumn.ColumnName + "]=");
                                sb.Append("@ARG" + index + ",");
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(" Where 1=1 ");
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            if (nodeTagColumn.PrimaryKey)
                            {
                                sb.Append(" AND [" + nodeTagColumn.ColumnName + "]=");
                                sb.Append("@ARG" + index + ",");
                                tabs.Add(nodeTagColumn);
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, false);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_UPDATE", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        private void ToolDataBase_QuerySel_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //SELECT TOP 100 [NO],[NAME] FROM [DBO].[TABLE1] WHERE [ID]=@ID
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("SELECT TOP(@ARG1)");
                        int index = 2;
                        List<NodeTag> tabs = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                sb.Append("[" + nodeTagColumn.ColumnName + "],");
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(" From [DBO].[" + nodeTag.TableName + "] ");
                        sb.Append(" Where 1=1 ");
                        List<NodeTag> templist = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                templist.Add(nodeTagColumn);
                            }
                        }
                        frm_database_select_mode dlg = new frm_database_select_mode();
                        dlg.Init(templist);
                        if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                        {
                            return;
                        }

                        index = 2;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                tabs.Add(nodeTagColumn); 
                                if (nodeTagColumn.QueryMode == NodeTag.eq)
                                {
                                    sb.Append(" AND [" + nodeTagColumn.ColumnName + "]=@ARG" + index);
                                }
                                else if (nodeTagColumn.QueryMode == NodeTag.Leftlike)
                                {
                                    sb.Append(" AND [" + nodeTagColumn.ColumnName + "] LIKE @ARG" + index + "+'%'");
                                }
                                else if (nodeTagColumn.QueryMode == NodeTag.Rightlike)
                                {
                                    sb.Append(" AND [" + nodeTagColumn.ColumnName + "] LIKE '%'+ @ARG" + index);
                                }
                                else if (nodeTagColumn.QueryMode == NodeTag.like)
                                {
                                    sb.Append(" AND [" + nodeTagColumn.ColumnName + "] LIKE '%'+@ARG" + index + "+'%'");
                                }
                                else
                                {
                                    sb.Append(" AND [" + nodeTagColumn.ColumnName + "]=@ARG" + index);
                                } 
                                //sb.Append(",");
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        index = 2;
                        sb.Append(" ORDER BY ");
                        int sortcount = 0;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (sortcount > 2)
                            {
                                break;
                            }
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                if (nodeTagColumn.IsDataTime)
                                {
                                    sb.Append(" [" + nodeTagColumn.ColumnName + "] DESC,");
                                }
                                else
                                {
                                    sb.Append(" [" + nodeTagColumn.ColumnName + "] ASC,");
                                }
                                sortcount++;
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, true);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_SELECT", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

        private void ToolDataBase_QueryEmptySel_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //SELECT TOP 100 [NO],[NAME] FROM [DBO].[TABLE1] WHERE {[ID]}=@ID
                List<string> list = new List<string>();
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("SELECT {@ARG1} ");
                        int index = 2;
                        List<NodeTag> tabs = new List<NodeTag>();
                        list.Add(" TOP (@ARG1)");
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                sb.Append("[" + nodeTagColumn.ColumnName + "],");
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(" From [DBO].[" + nodeTag.TableName + "] ");
                        sb.Append(" Where 1=1 ");
                        index = 2;
                        List<NodeTag> templist = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                templist.Add(nodeTagColumn);
                            }
                        }
                        frm_database_select_mode dlg = new frm_database_select_mode();
                        dlg.Init(templist);
                        if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                        {
                            return;
                        }

                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                tabs.Add(nodeTagColumn);
                                if (nodeTagColumn.QueryMode == NodeTag.eq)
                                {
                                    list.Add(" AND [" + nodeTagColumn.ColumnName + "]=@ARG" + index);
                                }
                                else if (nodeTagColumn.QueryMode == NodeTag.Leftlike)
                                {
                                    list.Add(" AND [" + nodeTagColumn.ColumnName + "] LIKE @ARG" + index + "+'%'");
                                }
                                else if (nodeTagColumn.QueryMode == NodeTag.Rightlike)
                                {
                                    list.Add(" AND [" + nodeTagColumn.ColumnName + "] LIKE '%'+ @ARG" + index);
                                }
                                else if (nodeTagColumn.QueryMode == NodeTag.like)
                                {
                                    list.Add(" AND [" + nodeTagColumn.ColumnName + "] LIKE '%'+@ARG" + index + "+'%'");
                                }
                                else
                                {
                                    list.Add(" AND [" + nodeTagColumn.ColumnName + "]=@ARG" + index);
                                }
                                sb.Append(" {@ARG" + index + "}");
                                index++;
                            }
                        }
                        //sb.Length = sb.Length - 1;
                        index = 2;
                        sb.Append(" ORDER BY ");
                        int sortcount = 0;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (sortcount > 1)
                            {
                                break;
                            }
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                if (nodeTagColumn.IsDataTime)
                                {
                                    sb.Append(" [" + nodeTagColumn.ColumnName + "] DESC,");
                                }
                                else
                                {
                                    sb.Append(" [" + nodeTagColumn.ColumnName + "] ASC,");
                                }
                                sortcount++;
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, true);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_SELECT_ARG", sql, code, list);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        private void ToolDataBase_QueryEmptyAll_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //SELECT TOP 100 [NO],[NAME] FROM [DBO].[TABLE1] WHERE {[ID]}=@ID
                List<string> list = new List<string>();
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("SELECT {@ARG1} ");
                        list.Add(" TOP(@ARG1)");
                        int index = 2;
                        List<NodeTag> tabs = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            sb.Append("[" + nodeTagColumn.ColumnName + "],");
                            index++;
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(" FROM [DBO].[" + nodeTag.TableName + "] ");
                        sb.Append(" Where 1=1 ");
                        index = 2;
                        List<NodeTag> templist = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                templist.Add(nodeTagColumn);
                            }
                        }
                        frm_database_select_mode dlg = new frm_database_select_mode();
                        dlg.Init(templist);
                        dlg.ShowDialog();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                tabs.Add(nodeTagColumn);
                                if (nodeTagColumn.QueryMode == NodeTag.eq)
                                {
                                    list.Add(" AND [" + nodeTagColumn.ColumnName + "]=@ARG" + index);
                                }
                                else if (nodeTagColumn.QueryMode == NodeTag.Leftlike)
                                {
                                    list.Add(" AND [" + nodeTagColumn.ColumnName + "] LIKE @ARG" + index + "+'%'");
                                }
                                else if (nodeTagColumn.QueryMode == NodeTag.Rightlike)
                                {
                                    list.Add(" AND [" + nodeTagColumn.ColumnName + "] LIKE '%'+ @ARG" + index);
                                }
                                else if (nodeTagColumn.QueryMode == NodeTag.like)
                                {
                                    list.Add(" AND [" + nodeTagColumn.ColumnName + "] LIKE '%'+@ARG" + index + "+'%'");
                                }
                                else
                                {
                                    list.Add(" AND [" + nodeTagColumn.ColumnName + "]=@ARG" + index);
                                }
                                sb.Append(" {@ARG" + index + "}");
                                index++;
                            }
                        }
                        index = 2;
                        sb.Append(" ORDER BY ");
                        int sortcount = 0;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (sortcount > 2)
                            {
                                break;
                            }
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                if (nodeTagColumn.IsDataTime)
                                {
                                    sb.Append(" [" + nodeTagColumn.ColumnName + "] DESC,");
                                }
                                else
                                {
                                    sb.Append(" [" + nodeTagColumn.ColumnName + "] ASC,");
                                }
                                sortcount++;
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, true);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_SELECT_ARG", sql, code, list);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

        private void ToolDataBase_QueryOneRowSel_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //SELECT  [NO],[NAME] FROM [DBO].[TABLE1] WHERE {[ID]}=@ID 
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        StringBuilder sbcode = new StringBuilder();
                        StringBuilder sbcodesetvalue = new StringBuilder();
                        sb.Append("SELECT TOP (1)");
                        int index = 2;
                        List<NodeTag> tabs = new List<NodeTag>();
                        sbcodesetvalue.Append("//");
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                sb.Append("[" + nodeTagColumn.ColumnName + "],");
                                sbcodesetvalue.Append(@"CELLVALUE(""" + nodeTagColumn.ColumnName + @""",DataTableRowValue(row,""" + nodeTagColumn.ColumnName + @"""));");
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(" From [DBO].[" + nodeTag.TableName + "] ");
                        sb.Append(" Where 1=1 ");
                        index = 2;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            if (nodeTagColumn != null)
                            {
                                if (nodeTagColumn.PrimaryKey)
                                {
                                    tabs.Add(nodeTagColumn);
                                    sb.Append(" AND [" + nodeTagColumn.ColumnName + "]=");
                                    sb.Append("@ARG" + index + ",");
                                    index++;
                                }
                            }
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, true, true);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_SELECT_Row", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }
        private void ToolDataBase_QueryOneRowAll_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //SELECT  [NO],[NAME] FROM [DBO].[TABLE1] WHERE {[ID]}=@ID 
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        StringBuilder sbcode = new StringBuilder();
                        StringBuilder sbcodesetvalue = new StringBuilder();
                        sb.Append("SELECT TOP (1)");
                        int index = 2;
                        List<NodeTag> tabs = new List<NodeTag>();
                        sbcodesetvalue.Append("//");
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            sb.Append("[" + nodeTagColumn.ColumnName + "],");
                            sbcodesetvalue.Append(@"CELLVALUE(""" + nodeTagColumn.ColumnName + @""",DataTableRowValue(row,""" + nodeTagColumn.ColumnName + @"""));");
                            index++;
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(" From [DBO].[" + nodeTag.TableName + "] ");
                        sb.Append(" Where 1=1 ");
                        index = 2;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            if (nodeTagColumn != null)
                            {
                                if (nodeTagColumn.PrimaryKey)
                                {
                                    tabs.Add(nodeTagColumn);
                                    sb.Append(" AND [" + nodeTagColumn.ColumnName + "]=");
                                    sb.Append("@ARG" + index + ",");
                                    index++;
                                }
                            }
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage(tabs, true, true);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_SELECT_Row", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

        private void ToolDataBase_CellValueSet_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //SELECT  [NO],[NAME] FROM [DBO].[TABLE1] WHERE {[ID]}=@ID 
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sbcodesetvalue = new StringBuilder();
                        int index = 2;
                        List<NodeTag> tabs = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                sbcodesetvalue.AppendLine(@"CELLVALUE(""" + nodeTagColumn.ColumnName + @""",DataTableRowValue(row,""" + nodeTagColumn.ColumnName + @"""));");
                                index++;
                            }
                        }
                        string code = sbcodesetvalue.ToString();
                        Feng.Forms.Dialogs.InputMultilineDialog dlg = new Forms.Dialogs.InputMultilineDialog();
                        dlg.Text = "附值代码";
                        dlg.txtInput.Text = code;
                        dlg.ShowDialog();
                        Feng.Forms.ClipboardHelper.SetClip(dlg.txtInput.Text);

                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

        private void ToolDataBase_EmptyVal_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //SELECT  [NO],[NAME] FROM [DBO].[TABLE1] WHERE {[ID]}=@ID 
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sbcodesetvalue = new StringBuilder();
                        int index = 2;
                        List<NodeTag> tabs = new List<NodeTag>();
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                sbcodesetvalue.Append(@"value=CELLVALUE(""" + nodeTagColumn.ColumnName + @""");
IF ISNULLOREMPTY(value)
    CellFocused(""" + nodeTagColumn.ColumnName + @""")
    RETURN 0;
ENDIF
");
                                index++;
                            }
                        }
                        string code = sbcodesetvalue.ToString();
                        Feng.Forms.Dialogs.InputMultilineDialog dlg = new Forms.Dialogs.InputMultilineDialog();
                        dlg.Text = "附值代码";
                        dlg.txtInput.Text = code;
                        dlg.ShowDialog();
                        Feng.Forms.ClipboardHelper.SetClip(dlg.txtInput.Text);

                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

        private void ToolDataBase_QueryExistSel_Click(object sender, EventArgs e)
        { 
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //SELECT  [NO],[NAME] FROM [DBO].[TABLE1] WHERE {[ID]}=@ID 
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        StringBuilder sbcode = new StringBuilder();
                        StringBuilder sbcodesetvalue = new StringBuilder();
                        sb.Append("SELECT TOP (1)");
                        int index = 2;
                        List<NodeTag> tabs = new List<NodeTag>();
                        sbcodesetvalue.Append("//");
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            if (item.Check)
                            {
                                NodeTag nodeTagColumn = item.Tag as NodeTag;
                                sb.Append("[" + nodeTagColumn.ColumnName + "],");
                                sbcodesetvalue.Append(@"CELLVALUE(""" + nodeTagColumn.ColumnName + @""",DataTableRowValue(row,""" + nodeTagColumn.ColumnName + @"""));");
                                index++;
                            }
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(" From [DBO].[" + nodeTag.TableName + "] ");
                        sb.Append(" Where 1=1 ");
                        index = 1;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            if (nodeTagColumn != null)
                            {
                                if (nodeTagColumn.PrimaryKey)
                                {
                                    tabs.Add(nodeTagColumn);
                                    sb.Append(" AND [" + nodeTagColumn.ColumnName + "]=");
                                    sb.Append("@ARG" + index + ",");
                                    index++;
                                }
                            }
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage_Exist(tabs);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_SELECT_exist", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

        private void ToolDataBase_QueryExistAll_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.Controls.TreeView.DataTreeNode node = this.treeviewcontroldatabase.TreeView.FocusedNode;
                if (node == null)
                    return;

                //SELECT  [NO],[NAME] FROM [DBO].[TABLE1] WHERE {[ID]}=@ID 
                NodeTag nodeTag = node.Tag as NodeTag;
                if (nodeTag != null)
                {
                    if (nodeTag.Type == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        StringBuilder sbcode = new StringBuilder();
                        StringBuilder sbcodesetvalue = new StringBuilder();
                        sb.Append("SELECT TOP (1)");
                        int index = 2;
                        List<NodeTag> tabs = new List<NodeTag>();
                        sbcodesetvalue.Append("//");
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            sb.Append("[" + nodeTagColumn.ColumnName + "],");
                            sbcodesetvalue.Append(@"CELLVALUE(""" + nodeTagColumn.ColumnName + @""",DataTableRowValue(row,""" + nodeTagColumn.ColumnName + @"""));");
                            index++;
                        }
                        sb.Length = sb.Length - 1;
                        sb.Append(" From [DBO].[" + nodeTag.TableName + "] ");
                        sb.Append(" Where 1=1 ");
                        index = 2;
                        foreach (Forms.Controls.TreeView.DataTreeNode item in node.Nodes)
                        {
                            NodeTag nodeTagColumn = item.Tag as NodeTag;
                            if (nodeTagColumn != null)
                            {
                                if (nodeTagColumn.PrimaryKey)
                                {
                                    tabs.Add(nodeTagColumn);
                                    sb.Append(" AND [" + nodeTagColumn.ColumnName + "]=");
                                    sb.Append("@ARG" + index + ",");
                                    index++;
                                }
                            }
                        }
                        sb.Length = sb.Length - 1;
                        string sql = sb.ToString();
                        string code = CreateScriptCodeNodeTage_Exist(tabs);
                        InsertCode(nodeTag.TableName, nodeTag.TableName + "_SELECT_exist", sql, code);
                    }

                }
            }
            catch (Exception ex)
            {
                Feng.Utils.TraceHelper.WriteTrace("DataDesign", "frmMain2", "statusSum_Click", ex);
            }
        }

    }
}
