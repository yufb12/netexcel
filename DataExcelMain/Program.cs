using Feng.Excel.App;
using Feng.Forms;
using System;
using System.Windows.Forms;
namespace Feng.DataDesign
{

    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [System.STAThread]
        static void Main(string[] args)
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            System.Windows.Forms.Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);

            try
            {

                if (args.Length > 0)
                {
                    if (System.IO.Path.GetExtension(args[0]) == ".xfm")
                    {
                        System.Windows.Forms.Application.Run(new frmForm(args[0]));
                    }
                    else
                    {
                        System.Windows.Forms.Application.Run(new frmMain2(args[0]));
                    }

                }
                else
                {
                    string file = Feng.IO.FileHelper.GetStartUpFileUSER("DataExcelMain", @"\Test" + Feng.App.FileExtension_DataExcel.DataExcel);
                    if (System.IO.File.Exists(file))
                    {
                        System.Windows.Forms.Application.Run(new frmMain2(file));
                    }
                    else
                    {
                        //using (SamllWaitingForm frm = new SamllWaitingForm())
                        //{
                        //    frm.Show();
                        //Feng.Forms.SplashForm.Start();
                        System.Windows.Forms.Application.Run(new frmMain2());

                        Feng.IO.LogHelper.Log("Exit", "10005");
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }


        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;
            if (ex != null)
            {
                Feng.Utils.ExceptionHelper.ShowError(ex);
            }
            else
            {
                Feng.Utils.ExceptionHelper.ShowError(e.ToString());
            }
        }

        static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            Feng.Utils.ExceptionHelper.ShowError(e.Exception);
        }
    }
}
