using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Feng.Excel
{
    public partial class frmFill : Form
    {
        public frmFill()
        {
            InitializeComponent();
        }

        private void btnok_Click(object sender, EventArgs e)
        {

            try
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
        
        }
    }
}
