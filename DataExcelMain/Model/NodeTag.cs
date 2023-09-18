using System;
using System.Data;

namespace Feng.Model
{
    public class NodeTag
    {
        public string Connection { get; set; }
        public string TableName { get; set; }
        public string ColumnName { get; set; }
        public string DataBase { get; set; }
        public int Type { get; set; }
        public DataRow Row { get; set; }
        public bool PrimaryKey { get; set; }
        public bool IsDataTime { get; set; }
        public bool IsInt { get; set; }
        public bool IsDecimal { get; set; }
        public bool IsString { get; set; }
        public bool IsBool { get; set; }
        public bool IDENTITY { get; set; }
        public string QueryMode { get; set; }
        public const string eq = "=";
        public const string like = "Like";
        public const string Leftlike = "Left Like";
        public const string Rightlike = "Right Like";
    }
}
 
