using Feng.Excel.Actions;
using Feng.Excel.Interfaces;
using System.Windows.Forms;

namespace Feng.Excel.Functions
{
  
    //public class PropertyActionTools
    //{
    //    public static System.Collections.Generic.List<CellPropertyAction> GetCellActions(ICell cell)
    //    {
            
    //        System.Collections.Generic.List<CellPropertyAction> list = new System.Collections.Generic.List<CellPropertyAction>();
    //        if (cell == null)
    //            return list;
    //        CellPropertyAction model = null;

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnCellEndEdit";
    //        //model.Script = cell.PropertyOnCellEndEdit;
    //        //model.Descript = "结束编辑";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnCellInitEdit";
    //        //model.Script = cell.PropertyOnCellInitEdit;
    //        //model.Descript = "初始化编辑";
    //        //list.Add(model);

    //        model = new CellPropertyAction();
    //        model.Cell = cell;
    //        model.ActionName = "PropertyOnCellValueChanged";
    //        model.Script = cell.PropertyOnCellValueChanged;
    //        model.Descript = "值变动";
    //        model.ShortName = "CellValueChanged";
    //        list.Add(model);

    //        model = new CellPropertyAction();
    //        model.Cell = cell;
    //        model.ActionName = "PropertyOnClick";
    //        model.Script = cell.PropertyOnClick;
    //        model.Descript = "单击";
    //        model.ShortName = "CellClick";
    //        list.Add(model);


    //        model = new CellPropertyAction();
    //        model.Cell = cell;
    //        model.ActionName = "PropertyOnDoubleClick";
    //        model.Script = cell.PropertyOnDoubleClick;
    //        model.Descript = "双击";
    //        model.ShortName = "CellDoubleClick";
    //        list.Add(model);


    //        model = new CellPropertyAction();
    //        model.Cell = cell;
    //        model.ActionName = "PropertyOnDrawBack";
    //        model.Script = cell.PropertyOnDrawBack;
    //        model.Descript = "单元格绘制背景";
    //        model.ShortName = "CellDrawBack";
    //        list.Add(model);


    //        model = new CellPropertyAction();
    //        model.Cell = cell;
    //        model.ActionName = "PropertyOnDrawCell";
    //        model.Script = cell.PropertyOnDrawCell;
    //        model.Descript = "单元格绘制前景";
    //        model.ShortName = "CellDraw";
    //        list.Add(model);


    //        model = new CellPropertyAction();
    //        model.Cell = cell;
    //        model.ActionName = "PropertyOnKeyDown";
    //        model.Script = cell.PropertyOnKeyDown;
    //        model.Descript = "键按下";
    //        model.ShortName = "CellKeyDown";
    //        list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnKeyPress";
    //        //model.Script = cell.PropertyOnKeyPress;
    //        //model.Descript = "键按下抬起";
    //        //list.Add(model);

    //        model = new CellPropertyAction();
    //        model.Cell = cell;
    //        model.ActionName = "PropertyOnKeyUp";
    //        model.Script = cell.PropertyOnKeyUp;
    //        model.Descript = "键抬起";
    //        model.ShortName = "CellKeyUp";
    //        list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnMouseCaptureChanged";
    //        //model.Script = cell.PropertyOnMouseCaptureChanged;
    //        //model.Descript = "捕捉到鼠标";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnMouseClick";
    //        //model.Script = cell.PropertyOnMouseClick;
    //        //model.Descript = "鼠标单击";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnMouseDoubleClick";
    //        //model.Script = cell.PropertyOnMouseDoubleClick;
    //        //model.Descript = "鼠标双击";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnMouseDown";
    //        //model.Script = cell.PropertyOnMouseDown;
    //        //model.Descript = "鼠标按下";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnMouseEnter";
    //        //model.Script = cell.PropertyOnMouseEnter;
    //        //model.Descript = "鼠标进入";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnMouseHover";
    //        //model.Script = cell.PropertyOnMouseHover;
    //        //model.Descript = "鼠标悬浮";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnMouseLeave";
    //        //model.Script = cell.PropertyOnMouseLeave;
    //        //model.Descript = "鼠标离开";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnMouseMove";
    //        //model.Script = cell.PropertyOnMouseMove;
    //        //model.Descript = "鼠标移动";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnMouseUp";
    //        //model.Script = cell.PropertyOnMouseUp;
    //        //model.Descript = "鼠标抬起";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnMouseWheel";
    //        //model.Script = cell.PropertyOnMouseWheel;
    //        //model.Descript = "鼠标滚轮滚动";
    //        //list.Add(model);

    //        //model = new CellPropertyAction();
    //        //model.Cell = cell;
    //        //model.ActionName = "PropertyOnPreviewKeyDown";
    //        //model.Script = cell.PropertyOnPreviewKeyDown;
    //        //model.Descript = "键按下前";
    //        //list.Add(model);

    //        return list;
    //    }

    //    public static void UpdateCellPropertyAction(CellPropertyAction model)
    //    {
    //        ICell cell = model.Cell;
    //        switch (model.ActionName)
    //        {
    //            case "PropertyOnCellEndEdit":
    //                cell.PropertyOnCellEndEdit = model.Script;
    //                break;
    //            case "PropertyOnCellInitEdit":
    //                cell.PropertyOnCellInitEdit = model.Script;
    //                break;
    //            case "PropertyOnCellValueChanged":
    //                cell.PropertyOnCellValueChanged = model.Script;
    //                break;
    //            case "PropertyOnClick":
    //                cell.PropertyOnClick = model.Script;
    //                break;
    //            case "PropertyOnDoubleClick":
    //                cell.PropertyOnDoubleClick = model.Script;
    //                break;
    //            case "PropertyOnDrawBack":
    //                cell.PropertyOnDrawBack = model.Script;
    //                break;
    //            case "PropertyOnDrawCell":
    //                cell.PropertyOnDrawCell = model.Script;
    //                break;
    //            case "PropertyOnKeyDown":
    //                cell.PropertyOnKeyDown = model.Script;
    //                break;
    //            case "PropertyOnKeyPress":
    //                cell.PropertyOnKeyPress = model.Script;
    //                break;
    //            case "PropertyOnKeyUp":
    //                cell.PropertyOnKeyUp = model.Script;
    //                break;
    //            case "PropertyOnMouseCaptureChanged":
    //                cell.PropertyOnMouseCaptureChanged = model.Script;
    //                break;
    //            case "PropertyOnMouseClick":
    //                cell.PropertyOnMouseClick = model.Script;
    //                break;
    //            case "PropertyOnMouseDoubleClick":
    //                cell.PropertyOnMouseDoubleClick = model.Script;
    //                break;
    //            case "PropertyOnMouseDown":
    //                cell.PropertyOnMouseDown = model.Script;
    //                break;
    //            case "PropertyOnMouseEnter":
    //                cell.PropertyOnMouseEnter = model.Script;
    //                break;
    //            case "PropertyOnMouseHover":
    //                cell.PropertyOnMouseHover = model.Script;
    //                break;
    //            case "PropertyOnMouseLeave":
    //                cell.PropertyOnMouseLeave = model.Script;
    //                break;
    //            case "PropertyOnMouseMove":
    //                cell.PropertyOnMouseMove = model.Script;
    //                break;
    //            case "PropertyOnMouseUp":
    //                cell.PropertyOnMouseUp = model.Script;
    //                break;
    //            case "PropertyOnMouseWheel":
    //                cell.PropertyOnMouseWheel = model.Script;
    //                break;
    //            case "PropertyOnPreviewKeyDown":
    //                cell.PropertyOnPreviewKeyDown = model.Script;
    //                break;
    //            default:
    //                break;
    //        }
    //    }

    //    public static System.Collections.Generic.List<DataExcelPropertyAction> GetGridActions(DataExcel grid)
    //    {
    //        System.Collections.Generic.List<DataExcelPropertyAction> list = new System.Collections.Generic.List<DataExcelPropertyAction>();
    //        DataExcelPropertyAction model = new DataExcelPropertyAction();
    //        model.Grid = grid;
    //        model.ActionName = "PropertyEndEdit";
    //        model.Script = grid.PropertyEndEdit;
    //        model.Descript = "结束编辑";
    //        model.ShortName = "EndEdit";
    //        list.Add(model); 
              

    //        model = new DataExcelPropertyAction();
    //        model.Grid = grid;
    //        model.ActionName = "PropertyFormClosing";
    //        model.Script = grid.PropertyFormClosing;
    //        model.Descript = "窗口关闭";
    //        model.ShortName = "FormClosing";
    //        list.Add(model); 

    //        model = new DataExcelPropertyAction();
    //        model.Grid = grid;
    //        model.ActionName = "PropertyLoadCompleted";
    //        model.Script = grid.PropertyDataLoadCompleted;
    //        model.Descript = "数据(文件)加载结束";
    //        model.ShortName = "LoadCompleted";
    //        list.Add(model);
             

    //        return list;
    //    }

    //    public static void UpdateDataExcelPropertyAction(DataExcel grid, DataExcelPropertyAction model)
    //    {
    //        switch (model.ActionName)
    //        {
    //            case "PropertyEndEdit":
    //                grid.PropertyEndEdit = model.Script;
    //                break;  
    //            case "PropertyFormClosing":
    //                grid.PropertyFormClosing = model.Script;
    //                break;  
    //            case "PropertyLoadCompleted":
    //                grid.PropertyDataLoadCompleted = model.Script;
    //                break; 
    //            default:
    //                break;
    //        }
    //    }
    //}
}