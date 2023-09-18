

using System.Windows.Forms;
using System;
using Feng.Forms;
using System.Collections.Generic;
using Feng.Excel.App;
using Feng.Excel.Interfaces;

namespace Feng.Excel
{

    public class CodeEdit : System.Windows.Forms.RichTextBox
    {
        private ICell cell = null;
        public ICell Cell { get { return cell; } }
        private EPerformEvent eperformevent =  EPerformEvent.Null;
        public EPerformEvent ePerformEvent { get { return eperformevent; } }
        public virtual void Format()
        {
            string text = this.Text;

        }
        public virtual void SetValue(ICell cel , EPerformEvent pe)
        {
            if (pe == EPerformEvent.Null)
                return;
            cell = cel;
            eperformevent = pe;
            SetValue();
        }
        private void SetValue()
        {
            switch (ePerformEvent)
            {
                case EPerformEvent.PropertyOnCellInitEdit:
                    this.Text = Cell.PropertyOnCellInitEdit;
                    break;
                case EPerformEvent.PropertyOnCellEndEdit:
                    this.Text = Cell.PropertyOnCellEndEdit;
                    break;
                case EPerformEvent.PropertyOnCellValueChanged:
                    this.Text = Cell.PropertyOnCellValueChanged;
                    break;
                case EPerformEvent.PropertyOnMouseUp:
                    this.Text = Cell.PropertyOnMouseUp;
                    break;
                case EPerformEvent.PropertyOnMouseMove:
                    this.Text = Cell.PropertyOnMouseMove;
                    break;
                case EPerformEvent.PropertyOnMouseLeave:
                    this.Text = Cell.PropertyOnMouseLeave;
                    break;
                case EPerformEvent.PropertyOnMouseHover:
                    this.Text = Cell.PropertyOnMouseHover;
                    break;
                case EPerformEvent.PropertyOnMouseEnter:
                    this.Text = Cell.PropertyOnMouseEnter;
                    break;
                case EPerformEvent.PropertyOnMouseDown:
                    this.Text = Cell.PropertyOnMouseDown;
                    break;
                case EPerformEvent.PropertyOnMouseDoubleClick:
                    this.Text = Cell.PropertyOnMouseDoubleClick;
                    break;
                case EPerformEvent.PropertyOnMouseClick:
                    this.Text = Cell.PropertyOnMouseClick;
                    break;
                case EPerformEvent.PropertyOnMouseCaptureChanged:
                    this.Text = Cell.PropertyOnMouseCaptureChanged;
                    break;
                case EPerformEvent.PropertyOnMouseWheel:
                    this.Text = Cell.PropertyOnMouseWheel;
                    break;
                case EPerformEvent.PropertyOnClick:
                    this.Text = Cell.PropertyOnClick;
                    break;
                case EPerformEvent.PropertyOnKeyDown:
                    this.Text = Cell.PropertyOnKeyDown;
                    break;
                case EPerformEvent.PropertyOnKeyPress:
                    this.Text = Cell.PropertyOnKeyPress;
                    break;
                case EPerformEvent.PropertyOnKeyUp:
                    this.Text = Cell.PropertyOnKeyUp;
                    break;
                case EPerformEvent.PropertyOnPreviewKeyDown:
                    this.Text = Cell.PropertyOnPreviewKeyDown;
                    break;
                case EPerformEvent.PropertyOnDoubleClick:
                    this.Text = Cell.PropertyOnDoubleClick;
                    break;
                case EPerformEvent.PropertyOnDrawBack:
                    this.Text = Cell.PropertyOnDrawBack;
                    break;
                case EPerformEvent.PropertyOnDrawCell:
                    this.Text = Cell.PropertyOnDrawCell;
                    break;
                default:
                    break;
            }
        }
        public virtual void UpdateValue(string value)
        {
            switch (ePerformEvent)
            {
                case EPerformEvent.PropertyOnCellInitEdit:
                    Cell.PropertyOnCellInitEdit = value;
                    break;
                case EPerformEvent.PropertyOnCellEndEdit:
                    Cell.PropertyOnCellEndEdit = value;
                    break;
                case EPerformEvent.PropertyOnCellValueChanged:
                    Cell.PropertyOnCellValueChanged = value;
                    break;
                case EPerformEvent.PropertyOnMouseUp:
                    Cell.PropertyOnMouseUp = value;
                    break;
                case EPerformEvent.PropertyOnMouseMove:
                    Cell.PropertyOnMouseMove = value;
                    break;
                case EPerformEvent.PropertyOnMouseLeave:
                    Cell.PropertyOnMouseLeave = value;
                    break;
                case EPerformEvent.PropertyOnMouseHover:
                    Cell.PropertyOnMouseHover = value;
                    break;
                case EPerformEvent.PropertyOnMouseEnter:
                    Cell.PropertyOnMouseEnter = value;
                    break;
                case EPerformEvent.PropertyOnMouseDown:
                    Cell.PropertyOnMouseDown = value;
                    break;
                case EPerformEvent.PropertyOnMouseDoubleClick:
                    Cell.PropertyOnMouseDoubleClick = value;
                    break;
                case EPerformEvent.PropertyOnMouseClick:
                    Cell.PropertyOnMouseClick = value;
                    break;
                case EPerformEvent.PropertyOnMouseCaptureChanged:
                    Cell.PropertyOnMouseCaptureChanged = value;
                    break;
                case EPerformEvent.PropertyOnMouseWheel:
                    Cell.PropertyOnMouseWheel = value;
                    break;
                case EPerformEvent.PropertyOnClick:
                    Cell.PropertyOnClick = value;
                    break;
                case EPerformEvent.PropertyOnKeyDown:
                    Cell.PropertyOnKeyDown = value;
                    break;
                case EPerformEvent.PropertyOnKeyPress:
                    Cell.PropertyOnKeyPress = value;
                    break;
                case EPerformEvent.PropertyOnKeyUp:
                    Cell.PropertyOnKeyUp = value;
                    break;
                case EPerformEvent.PropertyOnPreviewKeyDown:
                    Cell.PropertyOnPreviewKeyDown = value;
                    break;
                case EPerformEvent.PropertyOnDoubleClick:
                    Cell.PropertyOnDoubleClick = value;
                    break;
                case EPerformEvent.PropertyOnDrawBack:
                    Cell.PropertyOnDrawBack = value;
                    break;
                case EPerformEvent.PropertyOnDrawCell:
                    Cell.PropertyOnDrawCell = value;
                    break;
                default:
                    break;
            }
        }
        protected override void OnTextChanged(EventArgs e)
        {
            if (this.cell != null)
            {
                UpdateValue(this.cell.Text);
            }
            base.OnTextChanged(e);
        }
    }
    public enum EPerformEvent
    {
        Null,

        PropertyOnCellInitEdit,

        PropertyOnCellEndEdit,

        PropertyOnCellValueChanged,

        PropertyOnMouseUp,

        PropertyOnMouseMove,

        PropertyOnMouseLeave,

        PropertyOnMouseHover,

        PropertyOnMouseEnter,

        PropertyOnMouseDown,

        PropertyOnMouseDoubleClick,

        PropertyOnMouseClick,

        PropertyOnMouseCaptureChanged,

        PropertyOnMouseWheel,

        PropertyOnClick,

        PropertyOnKeyDown,

        PropertyOnKeyPress,

        PropertyOnKeyUp,

        PropertyOnPreviewKeyDown,

        PropertyOnDoubleClick,

        PropertyOnDrawBack,

        PropertyOnDrawCell
    }
}
