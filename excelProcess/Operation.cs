using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace ExcelOperation
{
    public class Operation
    {
        private static string _ExcelPath;  //Excel表格地址
        private static string _ConnStr;    //连接Excel表格字符串
        /// <summary>
        /// Excel表格地址
        /// </summary>
        public static string ExcelPath
        {
            get { return _ExcelPath; }
            set { _ExcelPath = value; }
        }
        /// <summary>
        /// 连接Excel表格字符串
        /// </summary>
        public static string ConnStr
        {
            get { return _ConnStr; }
            set { _ConnStr = value; }
        }
        /// <summary>
        /// 获取查询Excel表格结果
        /// </summary>
        /// <param name="strConn">连接Excel表字符串</param>
        /// <param name="ExcelSearchStr">Excel表格查询语句。注：from [Sheet1$]";//这里Sheet1对应excel的工作表名称，"$"符号是必须的</param>
        /// <returns>查询结果表</returns>
        public static DataTable GetExcelSearchTable(string strConn, string ExcelSearchStr)
        {
            DataSet ds = null;
            System.Data.DataTable dt = null;
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            OleDbDataAdapter myCommand = null;
            myCommand = new OleDbDataAdapter(ExcelSearchStr, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            dt = ds.Tables[0];
            conn.Close();
            return dt;
        }
    }
}
