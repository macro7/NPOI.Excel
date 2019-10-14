using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace ExcelHelperDemo
{
    /* ========================================================================
   * 【本类功能概述】   ExcelHelper Excel读取类
   * 
   * 作者：张宏杰       时间：2016/6/5 14:14:54
   * 文件名：WK_Framework.Utils.ExcelHelper
   * CLR版本：4.0.30319.235
   *
   * 修改者：           时间：              
   * 修改说明：
   * ========================================================================*/
    public class ExcelHelper {
        #region 读取Excel
        /// <summary>
        /// 读取Excel
        /// </summary>
        /// <param name="fileName">文件路径</param>
        /// <param name="sheetName">excel sheet 名称</param>
        /// <param name="isTitleOrDataOfFirstRow">if set to <c>true</c> [is title or data of first row].</param>
        /// <returns>DataSet.</returns>
        public static DataSet ExcelToDataSet(string fileName, string sheetName = "Sheet1", bool isTitleOrDataOfFirstRow = true) {
            //源的定义 
            if(string.IsNullOrEmpty(fileName)) {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.Multiselect = false;
                if(openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                    fileName = openFileDialog.FileName;
                }
                else {
                    return null;
                }
            }
            string fileType = System.IO.Path.GetExtension(fileName);
            string strConn = string.Format("Provider=Microsoft.Jet.OLEDB.{0}.0;" +
                        "Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";" +
                        "data source={3};",
                        (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (isTitleOrDataOfFirstRow ? "Yes" : "NO"), fileName);
            DataSet ds = null;
            using(OleDbConnection conn = new OleDbConnection(strConn)) {
                OleDbDataAdapter myCommand = null;
                try {
                    conn.Open();
                    string strExcel = "";
                    strExcel = "select * from [" + sheetName + "$]";
                    myCommand = new OleDbDataAdapter(strExcel, strConn);
                    ds = new DataSet();
                    myCommand.Fill(ds, "table1");
                }
                catch(SqlException ex) {
                    throw ex;
                }
                finally {
                    myCommand.Dispose();
                    conn.Close();
                }
                return ds;
            }
        }

        /// <summary>
        /// 将 Excel 文件转成 DataTable
        /// </summary>
        /// <param name="fileName">Excel文件及其路径</param>
        /// <param name="strSheetName">工作表名,如:Sheet1</param>
        /// <param name="isTitleOrDataOfFirstRow">True 第一行是标题,False 第一行是数据</param>
        /// <returns>DataTable</returns>
        public static DataTable ExcelToDataTable(string fileName, string sheetName = "Sheet1", bool isTitleOrDataOfFirstRow = true) {
            //源的定义 
            if(string.IsNullOrEmpty(fileName)) {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.Multiselect = false;
                if(openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                    fileName = openFileDialog.FileName;
                }
                else {
                    return null;
                }
            }
            string fileType = System.IO.Path.GetExtension(fileName);
            //源的定义 
            string strConn = string.Format("Provider=Microsoft.Jet.OLEDB.{0}.0;" +
                        "Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";" +
                        "data source={3};",
                        (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (isTitleOrDataOfFirstRow ? "Yes" : "NO"), fileName);
            //Sql语句
            //string strExcel = string.Format("select * from [{0}$]", strSheetName); 这是一种方法
            string strExcel = " SELECT * FROM [" + sheetName + "$]";
            //定义存放的数据表
            DataSet ds = new DataSet();

            //连接数据源
            using(OleDbConnection conn = new OleDbConnection(strConn)) {
                try {
                    conn.Open();
                    //适配到数据源
                    OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, strConn);

                    adapter.Fill(ds, sheetName);
                }
                catch(System.Data.SqlClient.SqlException ex) {
                    throw ex;
                }
                finally {
                    conn.Close();
                    conn.Dispose();
                }
            }
            return ds.Tables[sheetName];
        }
        #endregion

    }

}
