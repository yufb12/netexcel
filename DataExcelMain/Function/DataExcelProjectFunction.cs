using Feng.Excel.Interfaces;
using Feng.Script.CBEexpress;
using Feng.Script.Method;
using System.Windows.Forms;

namespace Feng.Excel.Functions
{
    public class DataProjectConfigMethodContainer : DataExcelMethodContainer
    {
        private static DataProjectConfigMethodContainer defaultmethod = null;
        public static DataProjectConfigMethodContainer DefaultMethod
        { 
            get {
                if (defaultmethod == null)
                {
                    defaultmethod = new DataProjectConfigMethodContainer();
                }
                return defaultmethod;
            }
        }
        public const string Function_Category = "Config";
        public const string Function_Description = "参数设置";
        public override string Name
        {
            get { return Function_Category; }

        }
        public override string Description
        {
            get { return Function_Description; }
        }
        private DataProjectConfigMethodContainer()
        {

            BaseMethod model = new BaseMethod();
            model.Name = "ConfigSave";
            model.Description = "参数设置 保存参数";
            model.Eg = @"ConfigSave(""NAME"",Cell(""A5""))";
            model.Function = ConfigSave;
            MethodList.Add(model);

            model = new BaseMethod();
            model.Name = "ConfigRead";
            model.Description = "参数设置 读取参数";
            model.Eg = @"ConfigRead(""NAME"",Cell(""A5""))";
            model.Function = ConfigRead;
            MethodList.Add(model);


            model = new BaseMethod();
            model.Name = "ConfigFileSave";
            model.Description = "参数设置 保存配置文件参数";
            model.Eg = @"ConfigFileSave(""/Config/User"",""NAME"",Cell(""A5""))";
            model.Function = ConfigFileSave;
            MethodList.Add(model);


            model = new BaseMethod();
            model.Name = "ConfigRead";
            model.Description = "参数设置 读取配置文件参数";
            model.Eg = @"ConfigRead(""NAME"",Cell(""A5""))";
            model.Function = ConfigRead;
            MethodList.Add(model);
        }

        private static Feng.Excel.DataExcel grid = null;
        public static Feng.Excel.DataExcel Grid
        {
            get
            {
                if (grid == null)
                {
                    grid = new DataExcel();
                    grid.Init();
                    string file = Feng.IO.FileHelper.GetStartUpFile("\\Config" + Feng.App.FileExtension_DataExcel.DataExcel);
                    if (System.IO.File.Exists(file))
                    {
                        grid.Open(file);
                    }
                    else
                    {
                        grid.FileName = file;
                    }
                }
                return grid;
            }
        }
        public object ConfigSave(params object[] args)
        {
            Feng.Excel.Script.DataExcelScriptStmtProxy value1 = base.GetArgIndex(0, args) as Feng.Excel.Script.DataExcelScriptStmtProxy;
            if (value1 == null)
                return Feng.Utils.Constants.FALSE;
            string name = base.GetTextValue(1, args);
            object value = base.GetValue(2, args);
            int columnindex = 2;
            for (int i = 1; i < 1000; i++)
            {
                ICell cell = Grid[i, columnindex];
                string key = Feng.Utils.ConvertHelper.ToString(cell.Value);
                if (string.IsNullOrWhiteSpace(key))
                {
                    cell.Value = name;
                }
                key = Feng.Utils.ConvertHelper.ToString(cell.Value);
                if (key == name)
                {
                    cell = Grid[i, columnindex + 1];
                    cell.Value = value;
                    break;
                }
            }
            Grid.Save();

            return Feng.Utils.Constants.TRUE;
        }

        public object ConfigRead(params object[] args)
        {
            Feng.Excel.Script.DataExcelScriptStmtProxy value1 = base.GetArgIndex(0, args) as Feng.Excel.Script.DataExcelScriptStmtProxy;
            if (value1 == null)
                return null;
            string name = base.GetTextValue(1, args);
            int columnindex = 2;
            for (int i = 1; i < 1000; i++)
            {
                try
                {
                    ICell cell = Grid[i, columnindex];
                    string key = Feng.Utils.ConvertHelper.ToString(cell.Value);
                    if (key == name)
                    {
                        cell = Grid[i, columnindex + 1];
                        return cell.Value;
                    }
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        return null;
                    }
                }
                catch (System.Exception)
                {
                    return null;
                }
            }
            return null;
        }

        public object ConfigFileSave(params object[] args)
        {
            Feng.Excel.Script.DataExcelScriptStmtProxy value1 = base.GetArgIndex(0, args) as Feng.Excel.Script.DataExcelScriptStmtProxy;
            if (value1 == null)
                return Feng.Utils.Constants.FALSE;
            string file = base.GetTextValue(1, args);
            string name = base.GetTextValue(2, args);
            object value = base.GetArgIndex(3, args);
            DataExcel grid = new DataExcel();
            grid.Init();
            if (!System.IO.Path.IsPathRooted(file))
            {
                file = Feng.IO.FileHelper.GetStartUpFile(file + Feng.App.FileExtension_DataExcel.DataExcel);
            }
            if (System.IO.File.Exists(file))
            {
                grid.Open(file);
            }
            else
            {
                grid.FileName = file;
            }
            int columnindex = 2;
            for (int i = 1; i < 1000; i++)
            {
                ICell cell = grid[i, columnindex];
                string key = Feng.Utils.ConvertHelper.ToString(cell.Value);
                if (string.IsNullOrWhiteSpace(key))
                {
                    cell.Value = name;
                }
                key = Feng.Utils.ConvertHelper.ToString(cell.Value);
                if (key == name)
                {
                    cell = grid[i, columnindex + 1];
                    cell.Value = value;
                    break;
                }
            }
            Grid.Save();

            return Feng.Utils.Constants.TRUE;
        }

        public object ConfigFileRead(params object[] args)
        {
            Feng.Excel.Script.DataExcelScriptStmtProxy value1 = base.GetArgIndex(0, args) as Feng.Excel.Script.DataExcelScriptStmtProxy;
            if (value1 == null)
                return Feng.Utils.Constants.FALSE;
            string file = base.GetTextValue(1, args);
            string name = base.GetTextValue(2, args);
            DataExcel grid = new DataExcel();
            grid.Init();
            if (!System.IO.Path.IsPathRooted(file))
            {
                file = Feng.IO.FileHelper.GetStartUpFile(file + Feng.App.FileExtension_DataExcel.DataExcel);
            }
            if (System.IO.File.Exists(file))
            {
                grid.Open(file);
            }
            else
            {
                grid.FileName = file;
            }
            int columnindex = 2;
            for (int i = 1; i < 1000; i++)
            { 
                ICell cell = grid.GetCell(i, columnindex);
                if (cell == null)
                    return null; 
                string key = Feng.Utils.ConvertHelper.ToString(cell.Value);
                if (key == name)
                {
                    cell = grid[i, columnindex + 1];
                    return cell.Value;
                }
            }
            return Feng.Utils.Constants.TRUE;
        }
    }
}