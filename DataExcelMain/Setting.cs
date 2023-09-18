using Feng.App;
using Feng.Excel.Interfaces;
using System;

namespace Feng.DataDesign
{
    public static class Fav
    {
        private static Feng.Excel.DataExcel fav = null;
        public static Feng.Excel.DataExcel FavCode
        {
            get
            {
                if (fav == null)
                {
                    fav = new Excel.DataExcel();
                    fav.Init();
                    string favcodefile = Feng.IO.FileHelper.GetStartUpFileUSER("DataExcelMain", "\\FavCode"+Feng.App.FileExtension_DataExcel .DataExcel);
                    if (System.IO.File.Exists(favcodefile))
                    {
                        fav.Open(favcodefile);
                    }
                    fav.FileName = favcodefile;
                }
                
                return fav;
            }
        }
        public static void AddFav(string name, string code)
        {
            for (int i = 1; i < 10000; i++)
            {
                ICell cell = FavCode[i, 2];
                if (string.IsNullOrEmpty(cell.Text))
                {
                    cell.Value = name;
                    cell = FavCode[i, 3];
                    cell.Value = code;
                    cell = FavCode[i, 4];
                    cell.Value = DateTime.Now;
                    break;
                }

            }
            FavCode.Save();
        }
    }

    public class Setting: Feng.IO.Setting
    {
        private static Setting setting = null;
        public static Setting Instance {
            get {
                if (setting == null)
                {
                    setting = new Setting();
                    setting.Init();
                }
                return setting;
            }
        }
        string file = Feng.IO.FileHelper .StartupPathUserData + "\\setting.set"; 
        public override string File
        { 
            get { return file; }  
            set { file = value; }
        }
 
        public string Mac
        {
            get
            {
                string mac = this.GetString("Mac");
                if (string.IsNullOrEmpty(mac))
                {
                    mac = DateTime.Now.ToString();
                    this.SetValue("Mac", mac);
                }
                return mac;
            }
            set
            {
                this.SetValue("Mac", value);
            }
        }

        public string Connecton
        {
            get
            {
                string value = this.GetString("Connecton");
                if (string.IsNullOrEmpty(value))
                    return string.Empty;
                string text = Feng.IO.DEncrypt.Decrypt(value, this.Mac);
                return text;
            }
            set
            {
                string text = Feng.IO.DEncrypt.Encrypt(value, this.Mac);
                this.SetValue("Connecton", text);
            }
        }


        public string Connecton1
        {
            get
            {
                string value = this.GetString("Connecton1");
                if (string.IsNullOrEmpty(value))
                    return string.Empty;
                string text = Feng.IO.DEncrypt.Decrypt(value, this.Mac);
                return text;
            }
            set
            {
                string text = Feng.IO.DEncrypt.Encrypt(value, this.Mac);
                this.SetValue("Connecton1", text);
            }
        }
        public string Connecton2
        {
            get
            {
                string value = this.GetString("Connecton2");
                if (string.IsNullOrEmpty(value))
                    return string.Empty;
                string text = Feng.IO.DEncrypt.Decrypt(value, this.Mac);
                return text;
            }
            set
            {
                string text = Feng.IO.DEncrypt.Encrypt(value, this.Mac);
                this.SetValue("Connecton2", text);
            }
        }
    }
}