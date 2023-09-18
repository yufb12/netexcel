using Feng.Excel;
using Feng.Excel.App;
using Feng.Excel.Args;
using Feng.Excel.Collections;
using Feng.Excel.Commands;
using Feng.Excel.Data;
using Feng.Excel.Delegates;
using Feng.Excel.Extend;
using Feng.Excel.Fillter;
using Feng.Excel.Interfaces;
using Feng.Forms;
using Feng.Forms.Command;
using Feng.Forms.Controls;
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
    public partial class frmMain2 : BaseForm
    {
        public void InitMenuTool()
        {
            InitBar();

            InitToolCommand();
            InitToolFunction();
            InitToolCommands();
            AddShowCommand();
        }
        private void InitBar()
        {
            InitDataBaseTool();
            InitIDTool();
            this.toolBarMainMenu.Skin.BarBackColor = this.BackColor;
            this.toolBarMainTool.Skin.BarBackColor = this.BackColor;
            this.toolBarMainMenu.BarItemHeader.Visable = false;
            this.toolBarMainMenu.ItemClick += ToolBarMainMenu_ItemClick;

            this.toolBarMainTool.ItemClick += ToolBarMainMenu_ItemClick;
            List<CommandObject> listtemp = new List<CommandObject>();
            Dictionary<string, Feng.Forms.Controls.ToolBarItem> listbaritem = new Dictionary<string, Feng.Forms.Controls.ToolBarItem>();
            Feng.Forms.Controls.ToolBarItem item = null;
            foreach (CommandObject command in this.dataexcel.CompositeKeys2.Commands)
            {
                if (string.IsNullOrWhiteSpace(command.ParentGroupName))
                {
                    if (!listbaritem.ContainsKey(command.GroupName))
                    {
                        item = new Feng.Forms.Controls.ToolBarItem(command.GroupTitle) { ID = command.GroupName, ShowImage = false, ShowToolTip = false };
                        this.toolBarMainMenu.Items.Add(item);
                        listbaritem.Add(command.GroupName, item);
                    }
                    item = listbaritem[command.GroupName];
                    item.Items.Add(new Feng.Forms.Controls.ToolBarItem(command.Description, command.CommandText,
                       GetImage(command.CommandText))
                    {
                        Tag = command,
                        MinWidth = 108,
                        ShowToolTip = false
                    });
                    listtemp.Add(command);
                }
            }
            for (int i = 0; i < 9; i++)
            {
                foreach (CommandObject command in this.dataexcel.CompositeKeys2.Commands)
                {
                    if (listtemp.Contains(command))
                    {
                        continue;
                    }
                    if (listbaritem.ContainsKey(command.ParentGroupName))
                    {
                        Feng.Forms.Controls.ToolBarItem itemparent = listbaritem[command.ParentGroupName];
                        if (!listbaritem.ContainsKey(command.GroupName))
                        {
                            item = new Feng.Forms.Controls.ToolBarItem(command.GroupTitle);
                            itemparent.Items.Add(item);
                            listbaritem.Add(command.GroupName, item);
                        }
                        item = listbaritem[command.GroupName];
                        string keytext = command.FirstKeyText + command.SencondKeyText;
                        int len = 50 - (command.Description + keytext).Length + command.Description.Length;
                        if (len < 1)
                        {
                            len = 0;
                        }
                        Feng.Forms.Controls.ToolBarItem toolBarItem = new Feng.Forms.Controls.ToolBarItem(command.Description + keytext, 
                            command.CommandText, GetImage(command.CommandText));
                        item.Items.Add(toolBarItem);
                        toolBarItem.Tag = command;
                        listtemp.Add(command);
                    }
                }
            }
        }
 
        public void InitImage()
        {
            Feng.Drawing.ImageCache.Add("Tool", "工具", CommandText.CommandSum, "合计", Feng.DataDesign.Properties.Resources.CommandSum);
            Feng.Drawing.ImageCache.Add("Tool", "工具", CommandText.CommandSumSelectCells, "合计选中", Feng.DataDesign.Properties.Resources.CommandSumSelectCells);
            Feng.Drawing.ImageCache.Add("Tool", "工具", CommandText.CommandToolFill, "填充", Feng.DataDesign.Properties.Resources.CommandSumSelectCells);
            Feng.Drawing.ImageCache.Add("Tool", "工具", CommandText.CommandToolLockRow, "锁定行", Feng.DataDesign.Properties.Resources.CommandSumSelectCells);
            Feng.Drawing.ImageCache.Add("Tool", "工具", CommandText.CommandToolUnLockRow, "解锁行", Feng.DataDesign.Properties.Resources.CommandSumSelectCells);

            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandCancel, "取消编辑", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandCopyAll, "复制", Feng.DataDesign.Properties.Resources.CommandCopyAll);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandCopyFormat, "复制样式", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandCopyID, "复制ID", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandCut, "剪切", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandDeleteCellMoveLeft, "删除单元格左移", Feng.DataDesign.Properties.Resources.CommandDeleteCellMoveLeft);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandDeleteCellMoveUp, "删除单元格上移", Feng.DataDesign.Properties.Resources.CommandDeleteCellMoveUp);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandDeleteColumn, "删除列", Feng.DataDesign.Properties.Resources.CommandDeleteColumn);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandDeleteRow, "删除行", Feng.DataDesign.Properties.Resources.CommandDeleteRow);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandFind, "查找", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandReplace, "替换", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandGo, "定位到", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandInsertCellMoveDown, "插入单元格下移", Feng.DataDesign.Properties.Resources.CommandInsertCellMoveDown);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandInsertCellMoveRight, "插入单元格右移", Feng.DataDesign.Properties.Resources.CommandInsertCellMoveRight);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandInsertColumn, "插入列", Feng.DataDesign.Properties.Resources.CommandInsertColumn);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandInsertRow, "插入行", Feng.DataDesign.Properties.Resources.CommandInsertRow);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandPaste, "粘贴", Feng.DataDesign.Properties.Resources.CommandEmpty); 
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandPasteBorder, "粘贴边框", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandPasteClear, "清除粘贴板", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandPasteFormat, "粘贴样式", Feng.DataDesign.Properties.Resources.CommandPasteFormat);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandPasteFormatBorder, "粘贴样式边框", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandPasteFormatColor, "粘贴颜色", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandPasteLoop, "循环粘贴", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandPasteText, "粘贴文本", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandRedo, "重做", Feng.DataDesign.Properties.Resources.CommandRedo);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandRemember, "开结/结束记忆命令", Feng.DataDesign.Properties.Resources.CommandRemember);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandRepeat, "重复命令", Feng.DataDesign.Properties.Resources.CommandRepeat);
            Feng.Drawing.ImageCache.Add("Edit", "编辑", CommandText.CommandUndo, "撤消", Feng.DataDesign.Properties.Resources.CommandUndo);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditCheckBox, "复选框编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditCheckBox);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellImageButton, "按钮控件", Feng.DataDesign.Properties.Resources.CommandCellImageButton);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditCnNumber, "中文大写数字金额", Feng.DataDesign.Properties.Resources.CommandCellEditCnNumber);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditComboBox, "下拉编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditComboBox);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellDropDownDateTime, "日期编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditDateTime);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditColor, "颜色编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditColor);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditGridView, "内嵌表格编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditGridView);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditImage, "图像编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditImage);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditLabel, "标签编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditLabel);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditLinkLabel, "链接编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditLinkLabel);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditNull, "清除编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditNull);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditNumber, "文本编辑控件(默认)", Feng.DataDesign.Properties.Resources.CommandCellEditNumber);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditPassword, "密码框编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditPassword);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditRadioBox, "单选框编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditRadioBox);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditTime, "时间编辑控件", Feng.DataDesign.Properties.Resources.CommandCellEditTime);
            Feng.Drawing.ImageCache.Add("EditControl", "编辑控件", CommandText.CommandCellEditTreeView, "树编辑控件", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("File", "文件", CommandText.CommandNew, "新建", Feng.DataDesign.Properties.Resources.CommandNew);
            Feng.Drawing.ImageCache.Add("File", "文件", CommandText.CommandOpen, "打开", Feng.DataDesign.Properties.Resources.CommandOpen);
            Feng.Drawing.ImageCache.Add("File", "文件", CommandText.CommandSave, "保存", Feng.DataDesign.Properties.Resources.CommandSave);
            Feng.Drawing.ImageCache.Add("File", "文件", CommandText.CommandSaveAs, "另保存", Feng.DataDesign.Properties.Resources.CommandSaveAs);
            Feng.Drawing.ImageCache.Add("File", "文件", CommandText.CommandPrint, "打印", Feng.DataDesign.Properties.Resources.CommandPrint);
            Feng.Drawing.ImageCache.Add("File", "文件", CommandText.CommandPrintView, "打印预览", Feng.DataDesign.Properties.Resources.CommandPrintView);
            Feng.Drawing.ImageCache.Add("File", "文件", CommandText.CommandPrintSetting, "打印设置", Feng.DataDesign.Properties.Resources.CommandPrintSetting);
            Feng.Drawing.ImageCache.Add("File", "文件", CommandText.CommandPrintArea, "打印设置", Feng.DataDesign.Properties.Resources.CommandPrintArea);

            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderBottom, "边框下边框", Feng.DataDesign.Properties.Resources.CommandBorderBottom);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderBottomClear, "边框取消下边框", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderClear, "边框清除", Feng.DataDesign.Properties.Resources.CommandBorderClear); 
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderFull, "边框全部", Feng.DataDesign.Properties.Resources.CommandBorderFull);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderLeft, "边框左边框", Feng.DataDesign.Properties.Resources.CommandBorderLeft);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderLeftBoomToRightTop, "边框左下至右上边框", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderLeftBottomToRightTopClear, "边框取消左下至右上边框", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderLeftClear, "边框取消左边框", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderLeftTopToRightBottom, "边框左上至右下边框", Feng.DataDesign.Properties.Resources.CommandBorderLeftTopToRightBottom);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderLeftTopToRightBottomClear, "边框取消左上至右下边框", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderRight, "边框右边框", Feng.DataDesign.Properties.Resources.CommandBorderRight);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderRightClear, "边框取消右边框", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderTop, "边框上边框", Feng.DataDesign.Properties.Resources.CommandBorderTop);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandBorderTopClear, "边框取消上边框", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellBackGround1, "背景颜色1", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellBackGround2, "背景颜色2", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellBackGround3, "背景颜色3", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellBackGround4, "背景颜色4", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellBackGround5, "背景颜色5", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellBackGround6, "背景颜色6", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellBackGround7, "背景颜色7", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellBackGroundDark, "背景颜色变深", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellBackGroundLight, "背景颜色变亮", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellReadOnly, "只读", Feng.DataDesign.Properties.Resources.CommandCellReadOnly);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandColumnAutoSize, "重置列宽", Feng.DataDesign.Properties.Resources.CommandColumnAutoSize);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandGridReadOnly, "单元格只读", Feng.DataDesign.Properties.Resources.CommandGridReadOnly);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandMergeCell, "合并/取消合并单元格", Feng.DataDesign.Properties.Resources.CommandMergeCell);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandMergeClear, "清除合并单元格", Feng.DataDesign.Properties.Resources.CommandEmpty);

            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellVisible, "内容可见", Feng.DataDesign.Properties.Resources.CommandCellVisable);
            Feng.Drawing.ImageCache.Add("Grid", "表格", CommandText.CommandCellHide, "内容不可见", Feng.DataDesign.Properties.Resources.CommandCellHide);

            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkAdd, "添加书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkDelete, "删除当前书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkEnd, "跳转到最后一个书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkFirst, "跳转到第一个书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkFooter, "跳转到最后一个书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkHeader, "跳转到首个书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkNext, "跳转到下个书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkNext, "跳转到下一个书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkNext, "跳转到下个书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkPrev, "跳转到上一个书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandBookmarkPrev, "跳转到上个书签", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandFirstCell, "移动到第一个单元格", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandFocusedCellNext, "下一个选中单元格", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandFocusedCellPrev, "上一个选中单元格", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandMoveFocusedCellToDown, "将当前焦点移动到下一个单元格", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandMoveFocusedCellToLeft, "将当前焦点移动到左边单元格", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandMoveFocusedCellToRight, "将当前焦点移动到右边单元格", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandMoveFocusedCellToTab, "将当前焦点移动到下个Tab", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandMoveFocusedCellToUp, "将当前焦点移动到上一个单元格", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectAll, "选择所有", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectDown, "向下选择", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectDownMove, "选中向下移动", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectDownText, "向下移动文本", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectLeft, "向左选择", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectLeftMove, "选中向左移动", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectLeftText, "向左移动文本", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectRight, "向右选择", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectRightMove, "选中向右移动", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectRightText, "向右移动文本", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectUp, "向上选择", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectUpMove, "选中向上移动", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Select", "选择", CommandText.CommandSelectUpText, "向上移动文本", Feng.DataDesign.Properties.Resources.CommandEmpty);

            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFont, "字体", Feng.DataDesign.Properties.Resources.CommandFont);

            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandCellBackGround, "背景色", Feng.DataDesign.Properties.Resources.CommandCellBackGround);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandCellBackImage, "背景图片", Feng.DataDesign.Properties.Resources.CommandCellBackImage);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandCellBackImageClear, "背景图片清除", Feng.DataDesign.Properties.Resources.CommandCellBackImageCancel);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandCellForeColor, "前景色", Feng.DataDesign.Properties.Resources.CommandCellForeColor);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontBold, "加粗字体", Feng.DataDesign.Properties.Resources.CommandFontBold);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontBoldCancel, "取消加粗字体", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontCancel, "重置默认字体", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontItalic, "斜体字体", Feng.DataDesign.Properties.Resources.CommandFontItalic);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontItalicCancel, "取消斜体字体", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontSizeDown, "字体大小减小", Feng.DataDesign.Properties.Resources.CommandFontSizeDown);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontSizeUp, "字体大小加大", Feng.DataDesign.Properties.Resources.CommandFontSizeUp);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontStrikeout, "删除线", Feng.DataDesign.Properties.Resources.CommandFontStrikeout);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontStrikeoutCancel, "取消删除线", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontUnderline, "下划线字体", Feng.DataDesign.Properties.Resources.CommandFontUnderline);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandFontUnderlineCancel, "取消下划线字体", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextAlignBottom, "文本下对齐", Feng.DataDesign.Properties.Resources.CommandTextAlignBottom);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextAlignCenter, "文本居中", Feng.DataDesign.Properties.Resources.CommandTextAlignCenter);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextAlignHorizontalCenter, "文本水平居中", Feng.DataDesign.Properties.Resources.CommandTextAlignHorizontalCenter);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextAlignVerticalCenter, "文本垂直居中", Feng.DataDesign.Properties.Resources.CommandTextAlignVerticalCenter);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextAlignLeft, "文本左对齐", Feng.DataDesign.Properties.Resources.CommandTextAlignLeft);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextAlignRight, "文本右对齐", Feng.DataDesign.Properties.Resources.CommandTextAlignRight);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextAlignTop, "文本上对齐", Feng.DataDesign.Properties.Resources.CommandTextAlignTop);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextOrientationRotateDown, "垂直文字", Feng.DataDesign.Properties.Resources.CommandTextOrientationRotateDown);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextAlignVerticalCenter, "文本垂直居中", Feng.DataDesign.Properties.Resources.CommandTextAlignVerticalCenter);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextAutoMultiline, "自动换行", Feng.DataDesign.Properties.Resources.CommandTextAutoMultiline);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextAutoMultilineCancel, "取消自动换行", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextTrimEndSpace, "去除文本尾部空格", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextTrimEndSymbol, "去除文本尾部符号", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextTrimSpace, "去除文本头尾空格", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextTrimStartSpace, "去除文本头部空格", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextTrimStartSymbol, "去除文本头部符号", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("Style", "样式", CommandText.CommandTextTrimSymbol, "去除文本头尾符号", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandGridShowColumnHeader, "显示列头", Feng.DataDesign.Properties.Resources.CommandGridShowColumnHeader);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandMulCellBackImage, "多单元格背景", Feng.DataDesign.Properties.Resources.CommandMulCellBackImage);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandMulCellBackImageCancel, "取消多单元格背景", Feng.DataDesign.Properties.Resources.CommandMulCellBackImageCancel);
            
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandGridShowGridColumnLine, "显示列表格线", Feng.DataDesign.Properties.Resources.CommandGridShowGridColumnLine);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandGridShowGridRowLine, "显示行表格线", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandGridShowHeader, "显示行列头", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandGridShowRowHeader, "显示行头", Feng.DataDesign.Properties.Resources.CommandGridShowRowHeader);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandHideCellInfo, "隐藏单元格信息", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandShowCellInfo, "显示单元格信息", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandShowHistory, "显示命令历史记录", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandShowRemember, "显示记忆命令", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandShowShortcut, "显示快捷键", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandShowRuler, "显示标尺", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandShowGridScroller, "显示滚动条", Feng.DataDesign.Properties.Resources.CommandEmpty);
            Feng.Drawing.ImageCache.Add("View", "视图", CommandText.CommandFrozen, "冻结到单元格", Feng.DataDesign.Properties.Resources.CommandFrozen);

            Feng.Drawing.ImageCache.Add("TextFormat", "文本格式", CommandText.CommandTextFormatMoney, "货币样式", Feng.DataDesign.Properties.Resources.CommandTextFormatMoney);
            Feng.Drawing.ImageCache.Add("TextFormat", "文本格式", CommandText.CommandTextFormatPercent, "百分比样式", Feng.DataDesign.Properties.Resources.CommandTextFormatPercent);
            Feng.Drawing.ImageCache.Add("TextFormat", "文本格式", CommandText.CommandTextFormatDecimalPlaces1, "只显示1位小数", Feng.DataDesign.Properties.Resources.CommandTextFormatDecimalPlaces1);
            Feng.Drawing.ImageCache.Add("TextFormat", "文本格式", CommandText.CommandTextFormatDecimalPlaces2, "只显示2位小数", Feng.DataDesign.Properties.Resources.CommandTextFormatDecimalPlaces2);
            Feng.Drawing.ImageCache.Add("TextFormat", "文本格式", CommandText.CommandTextFormatDateTimeDay, "日期样式", Feng.DataDesign.Properties.Resources.CommandTextFormatDateTimeDay);
            Feng.Drawing.ImageCache.Add("TextFormat", "文本格式", CommandText.CommandTextFormatDateTimeTime, "时间样式", Feng.DataDesign.Properties.Resources.CommandTextFormatDateTimeTime);
            Feng.Drawing.ImageCache.Add("TextFormat", "文本格式", CommandText.CommandTextFormatText, "时间样式", Feng.DataDesign.Properties.Resources.CommandTextFormatText);

        }
        public Image GetImage(string key)
        {
            Image img = Feng.Drawing.ImageCache.Get(key);
            if (img == null)
            {
                img = Feng.DataDesign.Properties.Resources.CommandEmpty;
            }
            return img;
        }
        public void AddToolCommand(string key)
        {
            if (key == CommandText.Split)
            {
                toolBarMainTool.Items.Add(new Feng.Forms.Controls.ToolBarItemVSplit());
            }
            else
            {
                CommandObject command = this.dataexcel.CompositeKeys2.GetCommand(key);
                if (command == null)
                    return;
                toolBarMainTool.Items.Add(new Feng.Forms.Controls.ToolBarItem(command.Description, command.CommandText,
          GetImage(command.CommandText), false, true)
                {
                    Tag = command,
                });
            }
        }
        public void AddShowCommand()
        {
            AddMenuCommand("View", new ExtendCommand()
            {
                CommandText = "ViewShowTree",
                CommandEvent = ShowTree,
                Description = "显示ID",
                Image = null
            });
            AddMenuCommand("View", new ExtendCommand()
            {
                CommandText = "ViewShowTree",
                CommandEvent = ShowPrintOut,
                Description = "显示事件",
                Image = null
            });
            AddMenuCommand("View", new ExtendCommand()
            {
                CommandText = "ViewShowTree",
                CommandEvent = ShowProperty,
                Description = "显示属性",
                Image = null
            });
            AddMenuCommand("Tool", new ExtendCommand()
            {
                CommandText = "FillCell",
                CommandEvent = FillCell,
                Description = "填充",
                Image = null
            });
            AddMenuCommand("Tool", new ExtendCommand()
            {
                CommandText = "LockRow",
                CommandEvent = LockRow,
                Description = "锁定行",
                Image = null
            });
            AddMenuCommand("Tool", new ExtendCommand()
            {
                CommandText = "UnLockRow",
                CommandEvent = UnLockRow,
                Description = "解锁行",
                Image = null
            });


            AddMenuCommand("Tool", new ExtendCommand()
            {
                CommandText = "FilterRow",
                CommandEvent = FilterRow,
                Description = "筛选",
                Image = null
            });


 
        }
        public void FilterRow(object sender, object value)
        {
            FilterExcel filterExcel = new FilterExcel();
            filterExcel.Init(this.dataexcel, this.dataexcel.SelectCells);
            this.dataexcel.FilterExcel = filterExcel;
        }
        public void LockRow(object sender, object value)
        {
            List<IRow> rows = this.dataexcel.GetSelectRows();
            foreach (IRow item in rows)
            {
                item.LockVersion = new Forms.Base.LockVersion() { };
            }
        }
        public void UnLockRow(object sender, object value)
        {
            List<IRow> rows = this.dataexcel.GetSelectRows();
            foreach (IRow item in rows)
            {
                item.LockVersion = null;
            }
        }
        public void ShowTree(object sender,object value)
        {
            splitContainerTreee.Panel1Collapsed = !splitContainerTreee.Panel1Collapsed;
            if (!splitContainerTreee.Panel1Collapsed) { this.RefreshID(); }
        }
        public void ShowProperty(object sender, object value)
        {
            panelMain.Panel2Collapsed = !panelMain.Panel2Collapsed;
        }
        public void ShowPrintOut(object sender, object value)
        {
            splitContainerPrintOut.Panel2Collapsed = !splitContainerPrintOut.Panel2Collapsed;
        }
        public void FillCell(object sender, object value)
        {
            ICell cell = this.dataexcel.FocusedCell;
            if (cell == null)
                return;
            using (frmFill frm = new frmFill())
            {
                frm.Icon = this.Icon;
                frm.StartPosition = FormStartPosition.CenterScreen;
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    int rowcount = Feng.Utils.ConvertHelper.ToInt32(frm.txtFillRowCount.Text);
                    int row = cell.Row.Index;
                    int column = cell.Column.Index;
                    for (int i = 1; i <= rowcount; i++)
                    {
                        ICell celltarget = this.dataexcel.GetCell(i + row, column);
                        if (frm.radioButtonAddNum.Checked)
                        {
                            decimal d = Feng.Utils.ConvertHelper.ToInt32(cell.Value);
                            celltarget.Value = d + Feng.Utils.ConvertHelper.ToDecimal(frm.txtAddNum.Text) * i;
                        }
                        if (frm.radioButtonAddTime.Checked)
                        {
                            DateTime d = Feng.Utils.ConvertHelper.ToDateTime(cell.Value, DateTime.Now.Date);
                            switch (frm.txtAddTimeUnit.Text)
                            {
                                case "天":
                                    celltarget.Value = d.AddDays( Feng.Utils.ConvertHelper.ToInt32(frm.txtAddTime.Text) * i);
                                    break;
                                case "时":
                                    celltarget.Value = d.AddHours(Feng.Utils.ConvertHelper.ToInt32(frm.txtAddTime.Text) * i);
                                    break;
                                case "分":
                                    celltarget.Value = d.AddMinutes(Feng.Utils.ConvertHelper.ToInt32(frm.txtAddTime.Text) * i);
                                    break;
                                case "秒":
                                    celltarget.Value = d.AddSeconds(Feng.Utils.ConvertHelper.ToInt32(frm.txtAddTime.Text) * i);
                                    break;
                                case "周":
                                    celltarget.Value = d.AddDays(Feng.Utils.ConvertHelper.ToInt32(frm.txtAddTime.Text) * i*7);
                                    break;
                                case "月":
                                    celltarget.Value = d.AddMonths(Feng.Utils.ConvertHelper.ToInt32(frm.txtAddTime.Text) * i * 7);
                                    break;
                                case "年":
                                    celltarget.Value = d.AddYears(Feng.Utils.ConvertHelper.ToInt32(frm.txtAddTime.Text) * i);
                                    break;
                                default:
                                    celltarget.Value = d.AddDays(Feng.Utils.ConvertHelper.ToInt32(frm.txtAddTime.Text) * i);
                                    break;
                            }
                        }
                        if (frm.radioButtonFixText.Checked)
                        {
                            celltarget.Value = frm.txtFixText.Text;
                        }
                        if (frm.radioButtonRandom.Checked)
                        {
                            celltarget.Value = Feng.Utils.RandomCache.Next(1000000,10000000);
                        }
                    }
                } 
            }
        }
        private void AddMenuCommand(string parentid, ExtendCommand command)
        {
            ToolBarItem item = this.toolBarMainMenu.Get(parentid);

            item.Items.Add(new Feng.Forms.Controls.ToolBarItem(command.Description, command.CommandText, command.Image, true, true)
            {
                Tag = command,
            }); 
        }
        private List<ExtendCommand> extendCommands = new List<ExtendCommand>();
        public void AddToolButton(ExtendCommand command)
        {
            ToolCommands.Add(command);
        }
        private void InitToolCommands()
        {
            foreach (ExtendCommand command in ToolCommands)
            {
                toolBarMainTool.Items.Add(new Feng.Forms.Controls.ToolBarItem(command.Description, command.CommandText ,command.Image, false, true)
                {
                    Tag = command,
                     ShowImage =true ,
                     ShowText=false
                });
            }
        }
        private void InitToolCommand()
        {
            AddToolCommand(CommandText.CommandSave);
            AddToolCommand(CommandText.CommandUndo);
            AddToolCommand(CommandText.CommandRedo);
            AddToolCommand(CommandText.CommandRemember);
            AddToolCommand(CommandText.CommandRepeat);
            AddToolCommand(CommandText.CommandFrozen);
            AddToolCommand(CommandText.CommandCopyAll);
            AddToolCommand(CommandText.CommandPasteFormat);
            AddToolCommand(CommandText.Split);

            AddToolCommand(CommandText.CommandTextAutoMultiline);
            AddToolCommand(CommandText.CommandColumnAutoSize);
            AddToolCommand(CommandText.CommandGridReadOnly); 
            AddToolCommand(CommandText.CommandMergeCell);
            AddToolCommand(CommandText.CommandCellHide);
            AddToolCommand(CommandText.CommandCellVisible);
            AddToolCommand(CommandText.Split);

            AddToolCommand(CommandText.CommandBorderFull);
            AddToolCommand(CommandText.CommandBorderRight);
            AddToolCommand(CommandText.CommandBorderLeft);
            AddToolCommand(CommandText.CommandBorderBottom);
            AddToolCommand(CommandText.CommandBorderTop);
            AddToolCommand(CommandText.CommandBorderClear);
            AddToolCommand(CommandText.Split);

            AddToolCommand(CommandText.CommandTextAlignCenter);
            AddToolCommand(CommandText.CommandTextAlignVerticalCenter);
            AddToolCommand(CommandText.CommandTextAlignHorizontalCenter);
            AddToolCommand(CommandText.CommandTextAlignLeft);
            AddToolCommand(CommandText.CommandTextAlignRight);
            AddToolCommand(CommandText.CommandTextAlignBottom);
            AddToolCommand(CommandText.CommandTextAlignTop);
            AddToolCommand(CommandText.CommandTextOrientationRotateDown);
            AddToolCommand(CommandText.Split);

            
            AddToolCommand(CommandText.CommandFont);
            AddToolCommand(CommandText.CommandFontSizeDown);
            AddToolCommand(CommandText.CommandFontSizeUp);
            AddToolCommand(CommandText.CommandFontUnderline);
            AddToolCommand(CommandText.CommandFontItalic);
            AddToolCommand(CommandText.CommandFontBold);
            AddToolCommand(CommandText.CommandCellBackGround);
            AddToolCommand(CommandText.CommandCellBackImage);
            AddToolCommand(CommandText.CommandCellBackImageClear);
            AddToolCommand(CommandText.CommandCellForeColor);

            AddToolCommand(CommandText.CommandMulCellBackImage);
            AddToolCommand(CommandText.CommandMulCellBackImageCancel);
            //AddToolCommand(CommandText.CommandFontCancel);
            AddToolCommand(CommandText.Split);

            AddToolCommand(CommandText.CommandInsertRow);
            AddToolCommand(CommandText.CommandInsertColumn);
            AddToolCommand(CommandText.CommandInsertCellMoveRight);
            AddToolCommand(CommandText.CommandInsertCellMoveDown);

            AddToolCommand(CommandText.CommandDeleteRow);
            AddToolCommand(CommandText.CommandDeleteColumn);
            AddToolCommand(CommandText.CommandDeleteCellMoveUp);
            AddToolCommand(CommandText.CommandDeleteCellMoveLeft);
            AddToolCommand(CommandText.Split);

            AddToolCommand(CommandText.CommandCellEditNull);
            AddToolCommand(CommandText.CommandCellImageButton);
            AddToolCommand(CommandText.CommandCellEditLabel);
            AddToolCommand(CommandText.CommandCellEditRadioBox);
            AddToolCommand(CommandText.CommandCellEditCheckBox);
            AddToolCommand(CommandText.CommandCellEditComboBox);
            AddToolCommand(CommandText.CommandCellEditPassword);
            AddToolCommand(CommandText.CommandCellEditNumber);
            AddToolCommand(CommandText.CommandCellEditColor);
            AddToolCommand(CommandText.CommandCellEditTime);
            AddToolCommand(CommandText.CommandCellEditImage);
            AddToolCommand(CommandText.CommandCellDropDownDateTime);
            AddToolCommand(CommandText.CommandCellEditCnNumber);
            AddToolCommand(CommandText.Split);

            AddToolCommand(CommandText.CommandTextFormatText);
            AddToolCommand(CommandText.CommandTextFormatMoney);
            AddToolCommand(CommandText.CommandTextFormatPercent);
            AddToolCommand(CommandText.CommandTextFormatDecimalPlaces1);
            AddToolCommand(CommandText.CommandTextFormatDecimalPlaces2);
            AddToolCommand(CommandText.CommandTextFormatDateTimeDay);
            AddToolCommand(CommandText.CommandTextFormatDateTimeTime);
            AddToolCommand(CommandText.Split);
        }
        public void InitToolFunction()
        {
            try
            {
                Feng.Forms.Controls.ToolBarItem item =
       new Feng.Forms.Controls.ToolBarItem("公式", null, true, false) { ShowToolTip = false };
                this.toolBarMainMenu.Items.Add(item);
                foreach (IMethod method in this.dataexcel.Methods)
                {
                    Feng.Forms.Controls.ToolBarItem toolmthod = new Feng.Forms.Controls.ToolBarItem(
                     method.Description,
                      method.Name,
                       GetImage(method.Description));
                    item.Items.Add(toolmthod);
                    toolmthod.Tag = method;
                    foreach (IMethodInfo funtion in method.MethodList)
                    {
                        Feng.Forms.Controls.ToolBarItem toolfun = new Feng.Forms.Controls.ToolBarItem(
                           funtion.Name + " " + funtion.Description,
                            funtion.Name,
                            GetImage(funtion.Name));
                        toolmthod.Items.Add(toolfun);
                        toolfun.Tag = funtion;
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.BugReport.Log(ex);
                Feng.Utils.TraceHelper.WriteTrace("DataExcelMain","frmMain2", "InitToolFunction",ex);
            }
            
        }
        private void LoasLastFileMenu()
        {
            //ToolStripLastFile.DropDownItems.Clear();
            for (int i = this.LastFiles.Count - 1; i >= 0; i--)
            {
                string file = this.LastFiles[i];
                //ToolStripItem item = ToolStripLastFile.DropDownItems.Add(file);
                //item.Click += new EventHandler(item_Click);
            }
        }
 
    }
}
