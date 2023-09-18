using System;
namespace Feng.Model
{
    /// <summary>
    /// 实体类clsemployee 。(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public class clsemployee
    {
        public clsemployee()
        { }
        #region Model
        private int _id;
        private string _emp_id;
        private string _fname;
        private string _minit;
        private string _lname;
        private int? _job_id;
        private int? _job_lvl;
        private string _pub_id;
        private DateTime? _hire_date;
        private decimal? _price;
        private decimal? _price1;
        /// <summary>
        ///字段名: id	 | 
        ///字段类型: int	 | 
        ///是否允许空: False	 | 
        ///长度: 4	 | 
        ///精度: 10	 | 
        ///小数位数: 0	 | 
        ///序号: 1	 | 
        ///默认值: 	 | 
        ///备注: 	
        ///是否是标识列: True	 | 
        ///是否是主键: True
        /// </summary>
        public int id
        {
            set { _id = value; }
            get { return _id; }
        }
        /// <summary>
        ///字段名: emp_id	 | 
        ///字段类型: varchar	 | 
        ///是否允许空: True	 | 
        ///长度: 50	 | 
        ///精度: 50	 | 
        ///小数位数: 0	 | 
        ///序号: 2	 | 
        ///默认值: 	 | 
        ///备注: 	
        ///是否是标识列: False	 | 
        ///是否是主键: False
        /// </summary>
        public string emp_id
        {
            set { _emp_id = value; }
            get { return _emp_id; }
        }
        /// <summary>
        ///字段名: fname	 | 
        ///字段类型: varchar	 | 
        ///是否允许空: True	 | 
        ///长度: 20	 | 
        ///精度: 20	 | 
        ///小数位数: 0	 | 
        ///序号: 3	 | 
        ///默认值: 	 | 
        ///备注: 	
        ///是否是标识列: False	 | 
        ///是否是主键: False
        /// </summary>
        public string fname
        {
            set { _fname = value; }
            get { return _fname; }
        }
        /// <summary>
        ///字段名: minit	 | 
        ///字段类型: varchar	 | 
        ///是否允许空: True	 | 
        ///长度: 50	 | 
        ///精度: 50	 | 
        ///小数位数: 0	 | 
        ///序号: 4	 | 
        ///默认值: 	 | 
        ///备注: 	
        ///是否是标识列: False	 | 
        ///是否是主键: False
        /// </summary>
        public string minit
        {
            set { _minit = value; }
            get { return _minit; }
        }
        /// <summary>
        ///字段名: lname	 | 
        ///字段类型: varchar	 | 
        ///是否允许空: True	 | 
        ///长度: 30	 | 
        ///精度: 30	 | 
        ///小数位数: 0	 | 
        ///序号: 5	 | 
        ///默认值: 	 | 
        ///备注: 	
        ///是否是标识列: False	 | 
        ///是否是主键: False
        /// </summary>
        public string lname
        {
            set { _lname = value; }
            get { return _lname; }
        }
        /// <summary>
        ///字段名: job_id	 | 
        ///字段类型: int	 | 
        ///是否允许空: True	 | 
        ///长度: 4	 | 
        ///精度: 10	 | 
        ///小数位数: 0	 | 
        ///序号: 6	 | 
        ///默认值: (1)	 | 
        ///备注: 	
        ///是否是标识列: False	 | 
        ///是否是主键: False
        /// </summary>
        public int? job_id
        {
            set { _job_id = value; }
            get { return _job_id; }
        }
        /// <summary>
        ///字段名: job_lvl	 | 
        ///字段类型: int	 | 
        ///是否允许空: True	 | 
        ///长度: 4	 | 
        ///精度: 10	 | 
        ///小数位数: 0	 | 
        ///序号: 7	 | 
        ///默认值: (10)	 | 
        ///备注: 	
        ///是否是标识列: False	 | 
        ///是否是主键: False
        /// </summary>
        public int? job_lvl
        {
            set { _job_lvl = value; }
            get { return _job_lvl; }
        }
        /// <summary>
        ///字段名: pub_id	 | 
        ///字段类型: varchar	 | 
        ///是否允许空: True	 | 
        ///长度: 50	 | 
        ///精度: 50	 | 
        ///小数位数: 0	 | 
        ///序号: 8	 | 
        ///默认值: ('9952')	 | 
        ///备注: 	
        ///是否是标识列: False	 | 
        ///是否是主键: False
        /// </summary>
        public string pub_id
        {
            set { _pub_id = value; }
            get { return _pub_id; }
        }
        /// <summary>
        ///字段名: hire_date	 | 
        ///字段类型: datetime	 | 
        ///是否允许空: True	 | 
        ///长度: 8	 | 
        ///精度: 23	 | 
        ///小数位数: 3	 | 
        ///序号: 9	 | 
        ///默认值: (getdate())	 | 
        ///备注: 	
        ///是否是标识列: False	 | 
        ///是否是主键: False
        /// </summary>
        public DateTime? hire_date
        {
            set { _hire_date = value; }
            get { return _hire_date; }
        }
        /// <summary>
        ///字段名: price	 | 
        ///字段类型: decimal	 | 
        ///是否允许空: True	 | 
        ///长度: 9	 | 
        ///精度: 18	 | 
        ///小数位数: 2	 | 
        ///序号: 10	 | 
        ///默认值: 	 | 
        ///备注: 	
        ///是否是标识列: False	 | 
        ///是否是主键: False
        /// </summary>
        public decimal? price
        {
            set { _price = value; }
            get { return _price; }
        }
        /// <summary>
        ///字段名: price1	 | 
        ///字段类型: decimal	 | 
        ///是否允许空: True	 | 
        ///长度: 9	 | 
        ///精度: 18	 | 
        ///小数位数: 2	 | 
        ///序号: 11	 | 
        ///默认值: 	 | 
        ///备注: 	
        ///是否是标识列: False	 | 
        ///是否是主键: False
        /// </summary>
        public decimal? price1
        {
            set { _price1 = value; }
            get { return _price1; }
        }
        #endregion Model

    }
}
 
