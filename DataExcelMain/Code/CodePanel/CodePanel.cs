

using System.Windows.Forms;
using System;
using Feng.Forms;
using System.Collections.Generic;
using Feng.Excel.App;
using Feng.Excel.Interfaces;

namespace Feng.Excel
{

    public class CodePanel : System.Windows.Forms.Panel
    {
        public ICell Cell { get; set; }
        
        public virtual void Format()
        {

        }
    }
}
