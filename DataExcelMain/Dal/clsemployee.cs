using System;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Data.SqlClient; 

using Feng.Data.MsSQL;
using Feng.Utils;
using Feng.Data;//请先添加引用
namespace Feng.DAL
{
    /// <summary>
    /// 数据访问类clsemployee。
    /// </summary>
    public class clsemployee
    {
        public clsemployee()
        { }
        #region  成员方法
        /// <summary>
        /// 是否存在该记录
        /// </summary>
        public bool Exists(int id)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(1) from employee with(nolock) ");
            strSql.Append(" where id=@id ");
            SqlParameter[] parameters = {
					new SqlParameter("@id", SqlDbType.Int,4)};
            parameters[0].Value = id;

            return Feng.Data.MsSQL.DbHelperSQL.Exists(strSql.ToString(), parameters);
        }


        /// <summary>
        /// 获取增加一条数据的信息
        /// </summary>
        public ModleInfo GetAddModelInfo(Feng.Model.clsemployee model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into employee(");
            strSql.Append("emp_id,fname,minit,lname,job_id,job_lvl,pub_id,hire_date,price,price1)");
            strSql.Append(" values (");
            strSql.Append("@emp_id,@fname,@minit,@lname,@job_id,@job_lvl,@pub_id,@hire_date,@price,@price1)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@emp_id", SqlDbType.VarChar,50),
					new SqlParameter("@fname", SqlDbType.VarChar,20),
					new SqlParameter("@minit", SqlDbType.VarChar,50),
					new SqlParameter("@lname", SqlDbType.VarChar,30),
					new SqlParameter("@job_id", SqlDbType.Int,4),
					new SqlParameter("@job_lvl", SqlDbType.Int,4),
					new SqlParameter("@pub_id", SqlDbType.VarChar,50),
					new SqlParameter("@hire_date", SqlDbType.DateTime),
					new SqlParameter("@price", SqlDbType.Decimal,9),
					new SqlParameter("@price1", SqlDbType.Decimal,9)};
            parameters[0].Value = model.emp_id;
            parameters[1].Value = model.fname;
            parameters[2].Value = model.minit;
            parameters[3].Value = model.lname;
            parameters[4].Value = model.job_id;
            parameters[5].Value = model.job_lvl;
            parameters[6].Value = model.pub_id;
            parameters[7].Value = model.hire_date;
            parameters[8].Value = model.price;
            parameters[9].Value = model.price1;

            return new ModleInfo(strSql.ToString(), parameters);
        }
        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Feng.Model.clsemployee model)
        {
            Feng.Data.ModleInfo mf = GetAddModelInfo(model);
            object obj = Feng.Data.MsSQL.DbHelperSQL.GetSingle(mf.Sql, mf.cmdParms);
            if (obj == null)
            {
                return 1;
            }
            else
            {
                return Convert.ToInt32(obj);
            }
        }
        /// <summary>
        /// 获取更新一条数据
        /// </summary>
        public ModleInfo GetUpdateModelInfo(Feng.Model.clsemployee model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update employee set ");
            strSql.Append("emp_id=@emp_id,");
            strSql.Append("fname=@fname,");
            strSql.Append("minit=@minit,");
            strSql.Append("lname=@lname,");
            strSql.Append("job_id=@job_id,");
            strSql.Append("job_lvl=@job_lvl,");
            strSql.Append("pub_id=@pub_id,");
            strSql.Append("hire_date=@hire_date,");
            strSql.Append("price=@price,");
            strSql.Append("price1=@price1");
            strSql.Append(" where id=@id ");
            SqlParameter[] parameters = {
					new SqlParameter("@id", SqlDbType.Int,4),
					new SqlParameter("@emp_id", SqlDbType.VarChar,50),
					new SqlParameter("@fname", SqlDbType.VarChar,20),
					new SqlParameter("@minit", SqlDbType.VarChar,50),
					new SqlParameter("@lname", SqlDbType.VarChar,30),
					new SqlParameter("@job_id", SqlDbType.Int,4),
					new SqlParameter("@job_lvl", SqlDbType.Int,4),
					new SqlParameter("@pub_id", SqlDbType.VarChar,50),
					new SqlParameter("@hire_date", SqlDbType.DateTime),
					new SqlParameter("@price", SqlDbType.Decimal,9),
					new SqlParameter("@price1", SqlDbType.Decimal,9)};
            parameters[0].Value = model.id;
            parameters[1].Value = model.emp_id;
            parameters[2].Value = model.fname;
            parameters[3].Value = model.minit;
            parameters[4].Value = model.lname;
            parameters[5].Value = model.job_id;
            parameters[6].Value = model.job_lvl;
            parameters[7].Value = model.pub_id;
            parameters[8].Value = model.hire_date;
            parameters[9].Value = model.price;
            parameters[10].Value = model.price1;

            return new ModleInfo(strSql.ToString(), parameters);
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public void Update(Feng.Model.clsemployee model)
        {
            ModleInfo mf = GetUpdateModelInfo(model);
            Feng.Data.MsSQL.DbHelperSQL.ExecuteSql(mf.Sql, mf.cmdParms);
        }

        /// <summary>
        /// 获取删除一条数据
        /// </summary>
        public ModleInfo GetDeleteModleInfo(int id)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete employee ");
            strSql.Append(" where id=@id ");
            SqlParameter[] parameters = {
					new SqlParameter("@id", SqlDbType.Int,4)};
            parameters[0].Value = id;

            return new ModleInfo(strSql.ToString(), parameters);
        }
        /// <summary>
        /// 删除一条数据
        /// </summary>
        public void Delete(int id)
        {
            ModleInfo mf = GetDeleteModleInfo(id);
            Feng.Data.MsSQL.DbHelperSQL.ExecuteSql(mf.Sql, mf.cmdParms);
        }
        /// <summary>
        /// 获取删除一条数据
        /// </summary>
        public ModleInfo GetDeleteDeleteInfo(DeleteInfo info)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete employee ");
            strSql.Append(" where " + info.Key + " = " + info.Value);
            SqlParameter[] parameters = { };
            return new ModleInfo(strSql.ToString(), parameters);
        }
        /// <summary>
        /// 删除一条数据
        /// </summary>
        public void Delete(DeleteInfo info)
        {
            ModleInfo mf = GetDeleteDeleteInfo(info);
            Feng.Data.MsSQL.DbHelperSQL.ExecuteSql(mf.Sql, mf.cmdParms);
        }


        /// <summary>
        /// 得到一个对象实体
        /// </summary>
        public Feng.Model.clsemployee GetModel(int id)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 id,emp_id,fname,minit,lname,job_id,job_lvl,pub_id,hire_date,price,price1 from employee with(nolock) ");
            strSql.Append(" where id=@id ");
            SqlParameter[] parameters = {
					new SqlParameter("@id", SqlDbType.Int,4)};
            parameters[0].Value = id;

            DataTable table = Feng.Data.MsSQL.DbHelperSQL.Query(strSql.ToString(), parameters);
            if (table.Rows.Count > 0)
            {
                System.Data.DataRow dr = table.Rows[0];
                return GetModel(dr);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 得到一个对象实体
        /// </summary>
        public Feng.Model.clsemployee GetModel(System.Data.DataRow dr)
        {
            Feng.Model.clsemployee model = new Feng.Model.clsemployee();
            if (dr != null)
            {
                if (dr["id"].ToString() != "")
                {
                    model.id = int.Parse(dr["id"].ToString());
                }
                model.emp_id = dr["emp_id"].ToString();
                model.fname = dr["fname"].ToString();
                model.minit = dr["minit"].ToString();
                model.lname = dr["lname"].ToString();
                if (dr["job_id"].ToString() != "")
                {
                    model.job_id = int.Parse(dr["job_id"].ToString());
                }
                if (dr["job_lvl"].ToString() != "")
                {
                    model.job_lvl = int.Parse(dr["job_lvl"].ToString());
                }
                model.pub_id = dr["pub_id"].ToString();
                if (dr["hire_date"].ToString() != "")
                {
                    model.hire_date = DateTime.Parse(dr["hire_date"].ToString());
                }
                if (dr["price"].ToString() != "")
                {
                    model.price = decimal.Parse(dr["price"].ToString());
                }
                if (dr["price1"].ToString() != "")
                {
                    model.price1 = decimal.Parse(dr["price1"].ToString());
                }
                return model;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 得到一个对象实体
        /// </summary>
        public Feng.Model.clsemployee GetModelByWhere(string where)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 id,emp_id,fname,minit,lname,job_id,job_lvl,pub_id,hire_date,price,price1 from employee with(nolock) ");
            if (where != string.Empty)
            {
                strSql.Append(" where " + where);
            }
            DataTable table = Feng.Data.MsSQL.DbHelperSQL.Query(strSql.ToString());
            if (table.Rows.Count > 0)
            {
                System.Data.DataRow dr = table.Rows[0];
                return GetModel(dr);
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// 得到对象实体列表
        /// </summary>
        public List<Feng.Model.clsemployee> GetModelList(string where)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select id,emp_id,fname,minit,lname,job_id,job_lvl,pub_id,hire_date,price,price1 from employee with(nolock) ");
            if (where != string.Empty)
            {
                strSql.Append(" where " + where);
            }
            DataTable table = Feng.Data.MsSQL.DbHelperSQL.Query(strSql.ToString());
            if (table.Rows.Count > 0)
            {
                return GetModelList(table);
            }
            else
            {
                return new List<Feng.Model.clsemployee>();
            }
        }


        /// <summary>
        /// 向数据集表中添加一行
        /// </summary>
        public void fnAddRow(System.Data.DataTable table, Feng.Model.clsemployee model, bool AcceptChanged)
        {
            System.Data.DataRow dr = table.NewRow();
            if (model != null)
            {
                dr["emp_id"] = model.emp_id;
                dr["fname"] = model.fname;
                dr["minit"] = model.minit;
                dr["lname"] = model.lname;
                dr["job_id"] = ConvertHelper.ToInt32(model.job_id);
                dr["job_lvl"] = ConvertHelper.ToInt32(model.job_lvl);
                dr["pub_id"] = model.pub_id;
                if (model.hire_date != null)
                {
                    dr["hire_date"] = model.hire_date;
                }
                dr["price"] = ConvertHelper.ToDecimal(model.price);
                dr["price1"] = ConvertHelper.ToDecimal(model.price1);
                table.Rows.Add(dr);
                if (AcceptChanged)
                {
                    table.AcceptChanges();
                }
            }
        }

        /// <summary>
        /// 更新表中的一行
        /// </summary>
        public void fnSetRow(System.Data.DataRow dr, Feng.Model.clsemployee model)
        {
            if (model != null)
            {
                dr["emp_id"] = model.emp_id;
                dr["fname"] = model.fname;
                dr["minit"] = model.minit;
                dr["lname"] = model.lname;
                dr["job_id"] = ConvertHelper.ToInt32(model.job_id);
                dr["job_lvl"] = ConvertHelper.ToInt32(model.job_lvl);
                dr["pub_id"] = model.pub_id;
                if (model.hire_date != null)
                {
                    dr["hire_date"] = model.hire_date;
                }
                dr["price"] = ConvertHelper.ToDecimal(model.price);
                dr["price1"] = ConvertHelper.ToDecimal(model.price1);
            }
        }

        /// <summary>
        /// 获得数据列表
        /// </summary>
        public DataTable GetList(string strWhere)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select id,emp_id,fname,minit,lname,job_id,job_lvl,pub_id,hire_date,price,price1 ");
            strSql.Append(" FROM employee with(nolock) ");
            if (strWhere.Trim() != "")
            {
                strSql.Append(" where " + strWhere);
            }
            return Feng.Data.MsSQL.DbHelperSQL.Query(strSql.ToString());
        }
        /// <summary>
        /// 获得数据列表
        /// </summary>
        public DataTable GetList(string strWhere, string strselect)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select " + strselect);
            strSql.Append(" FROM employee with(nolock) ");
            if (strWhere.Trim() != "")
            {
                strSql.Append(" where " + strWhere);
            }
            return Feng.Data.MsSQL.DbHelperSQL.Query(strSql.ToString());
        }
        /// <summary>
        /// 获得数据列表
        /// </summary>
        public DataTable GetList(string strWhere, StringBuilder sbfield, StringBuilder sbon)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append(@"select 				T.id,
				T.emp_id,
				T.fname,
				T.minit,
				T.lname,
				T.job_id,
				T.job_lvl,
				T.pub_id,
				T.hire_date,
				T.price,
				T.price1 ");
            if (sbfield.ToString() != string.Empty)
            {
                strSql.Append("," + sbfield.ToString());
            }
            strSql.Append(" FROM employee AS T  WITH (NOLOCK) ");
            if (sbon.ToString() != string.Empty)
            {
                sbon.ToString();
            }
            if (strWhere.Trim() != "")
            {
                strSql.Append(" where " + strWhere);
            }
            return Feng.Data.MsSQL.DbHelperSQL.Query(strSql.ToString());
        }
        /// <summary>
        /// 获得数据列表
        /// </summary>
        public DataTable GetList(StringBuilder strSql)
        {
            return Feng.Data.MsSQL.DbHelperSQL.Query(strSql.ToString());
        }

        /// <summary>
        /// 获得前几行数据
        /// </summary>
        public DataTable GetList(int Top, string strWhere, string filedOrder)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select ");
            if (Top > 0)
            {
                strSql.Append(" top " + Top.ToString());
            }
            strSql.Append(" id,emp_id,fname,minit,lname,job_id,job_lvl,pub_id,hire_date,price,price1 ");
            strSql.Append(" FROM employee with(nolock) ");
            if (strWhere.Trim() != "")
            {
                strSql.Append(" where " + strWhere);
            }
            if (filedOrder != string.Empty)
            {
                strSql.Append(" order by " + filedOrder);
            }
            return Feng.Data.MsSQL.DbHelperSQL.Query(strSql.ToString());
        }

        /// <summary>
        /// 获得数据列表
        /// </summary>
        public int fnQuery(string strWhere, DataTable table, bool AutoClear)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select id,emp_id,fname,minit,lname,job_id,job_lvl,pub_id,hire_date,price,price1 ");
            strSql.Append(" FROM employee with(nolock) ");
            if (strWhere.Trim() != "")
            {
                strSql.Append(" where " + strWhere);
            }
            return fnGetList(strSql.ToString(), table, AutoClear);
        }
        /// <summary>
        /// 获得数据列表
        /// </summary>
        public int fnQuery(string strWhere, string strselect, DataTable table, bool AutoClear)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select " + strselect);
            strSql.Append(" FROM employee with(nolock) ");
            if (strWhere.Trim() != "")
            {
                strSql.Append(" where " + strWhere);
            }
            return fnGetList(strSql.ToString(), table, AutoClear);
        }
        /// <summary>
        /// 获得数据列表
        /// </summary>
        public int fnQuery(string strWhere, StringBuilder sbfield, StringBuilder sbon, DataTable table, bool AutoClear)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append(@"select 				T.id,
				T.emp_id,
				T.fname,
				T.minit,
				T.lname,
				T.job_id,
				T.job_lvl,
				T.pub_id,
				T.hire_date,
				T.price,
				T.price1 ");
            if (sbfield.ToString() != string.Empty)
            {
                strSql.Append("," + sbfield.ToString());
            }
            strSql.Append(" FROM employee AS T  WITH (NOLOCK) ");
            if (sbon.ToString() != string.Empty)
            {
                sbon.ToString();
            }
            if (strWhere.Trim() != "")
            {
                strSql.Append(" where " + strWhere);
            }
            return fnGetList(strSql.ToString(), table, AutoClear);
        }
        /// <summary>
        /// 获得数据列表
        /// </summary>
        public int fnQuery(StringBuilder strSql, DataTable table, bool AutoClear)
        {
            return fnGetList(strSql.ToString(), table, AutoClear);
        }

        /// <summary>
        /// 获得前几行数据
        /// </summary>
        public int fnQuery(int Top, string strWhere, string filedOrder, DataTable table, bool AutoClear)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select ");
            if (Top > 0)
            {
                strSql.Append(" top " + Top.ToString());
            }
            strSql.Append(" id,emp_id,fname,minit,lname,job_id,job_lvl,pub_id,hire_date,price,price1 ");
            strSql.Append(" FROM employee with(nolock) ");
            if (strWhere.Trim() != "")
            {
                strSql.Append(" where " + strWhere);
            }
            if (filedOrder != string.Empty)
            {
                strSql.Append(" order by " + filedOrder);
            }
            return fnGetList(strSql.ToString(), table, AutoClear);
        }

        /// <summary>
        /// 通用查询语句
        /// </summary>
        public int fnGetList(string sqlstr, DataTable table, bool AutoClear)
        {
            DataTable ds = Feng.Data.MsSQL.DbHelperSQL.Query(sqlstr);
 
            if (AutoClear)
            {
                table.Clear();
            }
            table.Merge(ds);
            return table.Rows.Count;
        }

        /// <summary>
        /// 得到一个对象实体
        /// </summary>
        public List<Feng.Model.clsemployee> GetModelList(DataTable table)
        {
            List<Feng.Model.clsemployee> list = new List<Feng.Model.clsemployee>();
            foreach (DataRow row in table.Rows)
            {
                list.Add(GetModel(row));
            }
            return list;
        }

        /*
        /// <summary>
        /// 分页获取数据列表
        /// </summary>
        public DataSet GetList(int PageSize,int PageIndex,string strWhere)
        {
            SqlParameter[] parameters = {
                    new SqlParameter("@tblName", SqlDbType.VarChar, 255),
                    new SqlParameter("@fldName", SqlDbType.VarChar, 255),
                    new SqlParameter("@PageSize", SqlDbType.Int),
                    new SqlParameter("@PageIndex", SqlDbType.Int),
                    new SqlParameter("@IsReCount", SqlDbType.Bit),
                    new SqlParameter("@OrderType", SqlDbType.Bit),
                    new SqlParameter("@strWhere", SqlDbType.VarChar,1000),
                    };
            parameters[0].Value = "employee";
            parameters[1].Value = "ID";
            parameters[2].Value = PageSize;
            parameters[3].Value = PageIndex;
            parameters[4].Value = 0;
            parameters[5].Value = 0;
            parameters[6].Value = strWhere;	
            return Feng.Data.MsSQL.MsMsDbHelperSQL.RunProcedure("UP_GetRecordByPage",parameters,"ds");
        }*/

        #endregion  成员方法
    }
}

