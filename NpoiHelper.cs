using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;

namespace ExcelHelperDemo
{
    /// <summary>
       /// Excel文件到DataSet的转换类
       /// </summary>
    public class NpoiHelper
    {
        #region 读取Excel文件内容转换为DataSet
        /// <summary>
               /// 读取Excel文件内容转换为DataSet,列名依次为 "c0"……c[columnlength-1]
               /// </summary>
               /// <param name="FileName">文件绝对路径</param>
               /// <param name="startRow">数据开始行数(1为第一行)</param>
               /// <param name="ColumnDataType">每列的数据类型</param>
               /// <returns></returns>
        public DataSet ReadExcel(string FileName, int startRow, params NpoiDataType[] ColumnDataType)
        {
            int ertime = 0;
            int intime = 0;
            DataSet ds = new DataSet("ds");
            DataTable dt = new DataTable("dt");
            DataRow dr;
            StringBuilder sb = new StringBuilder();
            using (FileStream stream = new FileStream(@FileName, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = WorkbookFactory.Create(stream);//使用接口，自动识别excel2003/2007格式
                ISheet sheet = workbook.GetSheetAt(0);//得到里面第一个sheet
                int j;
                IRow row;
                #region ColumnDataType赋值
                if (ColumnDataType.Length <= 0)
                {
                    row = sheet.GetRow(startRow - 1);//得到第i行
                    ColumnDataType = new NpoiDataType[row.LastCellNum];
                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        ICell hs = row.GetCell(i);
                        ColumnDataType[i] = GetCellDataType(hs);
                    }
                }
                #endregion
                for (j = 0; j < ColumnDataType.Length; j++)
                {
                    Type tp = GetDataTableType(ColumnDataType[j]);
                    dt.Columns.Add("c" + j, tp);
                }
                for (int i = startRow - 1; i <= sheet.PhysicalNumberOfRows; i++)
                {
                    row = sheet.GetRow(i);//得到第i行
                    if (row == null)
                    {
                        continue;
                    }

                    try
                    {
                        dr = dt.NewRow();

                        for (j = 0; j < ColumnDataType.Length; j++)
                        {
                            dr["c" + j] = GetCellData(ColumnDataType[j], row, j);
                        }
                        dt.Rows.Add(dr);
                        intime++;
                    }
                    catch (Exception er)
                    {
                        ertime++;
                        sb.Append(string.Format("第{0}行出错：{1}\r\n", i + 1, er.Message));
                        continue;
                    }
                }
                ds.Tables.Add(dt);
            }
            if (ds.Tables[0].Rows.Count == 0 && sb.ToString() != "")
            {
                throw new Exception(sb.ToString());
            }

            return ds;
        }
        #endregion
        Color LevelOneColor = Color.Green;
        Color LevelTwoColor = Color.FromArgb(201, 217, 243);
        Color LevelThreeColor = Color.FromArgb(231, 238, 248);
        Color LevelFourColor = Color.FromArgb(232, 230, 231);
        Color LevelFiveColor = Color.FromArgb(250, 252, 213);

        #region 从DataSet导出到MemoryStream流2003
        /// <summary>
               /// 从DataSet导出到MemoryStream流2003
               /// </summary>
               /// <param name="SaveFileName">文件保存路径</param>
               /// <param name="SheetName">Excel文件中的Sheet名称</param>
               /// <param name="ds">存储数据的DataSet</param>
               /// <param name="startRow">从哪一行开始写入，从0开始</param>
               /// <param name="datatypes">DataSet中的各列对应的数据类型</param>
        public bool CreateExcel2003(string SaveFileName, string SheetName, DataSet ds, int startRow, params NpoiDataType[] datatypes)
        {
            try
            {
                if (startRow < 0)
                {
                    startRow = 0;
                }

                HSSFWorkbook wb = new HSSFWorkbook();
                wb = new HSSFWorkbook();
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "pkm";
                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Title =
                si.Subject = "automatic genereted document";
                si.Author = "pkm";
                wb.DocumentSummaryInformation = dsi;
                wb.SummaryInformation = si;
                ISheet sheet = wb.CreateSheet(SheetName);
                //sheet.SetColumnWidth(0, 50 * 256);
                //sheet.SetColumnWidth(1, 100 * 256);
                IRow row;
                ICell cell;
                DataRow dr;
                int j;
                int maxLength = 0;
                int curLength = 0;
                object columnValue;
                DataTable dt = ds.Tables[0];
                if (datatypes.Length < dt.Columns.Count)
                {
                    datatypes = new NpoiDataType[dt.Columns.Count];
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        string dtcolumntype = dt.Columns[i].DataType.Name.ToLower();
                        switch (dtcolumntype)
                        {
                            case "string":
                                datatypes[i] = NpoiDataType.String;
                                break;
                            case "datetime":
                                datatypes[i] = NpoiDataType.Datetime;
                                break;
                            case "boolean":
                                datatypes[i] = NpoiDataType.Bool;
                                break;
                            case "double":
                                datatypes[i] = NpoiDataType.Numeric;
                                break;
                            default:
                                datatypes[i] = NpoiDataType.String;
                                break;
                        }
                    }
                }

                #region 创建表头
                row = sheet.CreateRow(0);//创建第i行
                ICellStyle style1 = wb.CreateCellStyle();//样式
                IFont font1 = wb.CreateFont();//字体

                font1.Color = HSSFColor.White.Index;//字体颜色
                font1.Boldweight = (short)FontBoldWeight.Bold;//字体加粗样式
                                                              //style1.FillBackgroundColor = HSSFColor.WHITE.index;//GetXLColour(wb, LevelOneColor);// 设置图案色
                style1.FillForegroundColor = HSSFColor.Green.Index;//GetXLColour(wb, LevelOneColor);// 设置背景色
                style1.FillPattern = FillPattern.SolidForeground;
                style1.SetFont(font1);//样式里的字体设置具体的字体样式
                style1.Alignment = HorizontalAlignment.Center;//文字水平对齐方式
                style1.VerticalAlignment = VerticalAlignment.Center;//文字垂直对齐方式
                row.HeightInPoints = 25;
                for (j = 0; j < dt.Columns.Count; j++)
                {
                    columnValue = dt.Columns[j].ColumnName;
                    curLength = Encoding.Default.GetByteCount(columnValue.ToString());
                    maxLength = (maxLength < curLength ? curLength : maxLength);
                    int colounwidth = 256 * maxLength;
                    sheet.SetColumnWidth(j, colounwidth);
                    try
                    {
                        cell = row.CreateCell(j);//创建第0行的第j列
                        cell.CellStyle = style1;//单元格式设置样式

                        try
                        {
                            cell.SetCellType(CellType.String);
                            cell.SetCellValue(columnValue.ToString());
                        }
                        catch { }

                    }
                    catch
                    {
                        continue;
                    }
                }
                #endregion

                #region 创建每一行
                for (int i = startRow; i < ds.Tables[0].Rows.Count; i++)
                {
                    dr = ds.Tables[0].Rows[i];
                    row = sheet.CreateRow(i + 1);//创建第i行
                    for (j = 0; j < dt.Columns.Count; j++)
                    {
                        columnValue = dr[j];
                        curLength = Encoding.Default.GetByteCount(columnValue.ToString());
                        maxLength = (maxLength < curLength ? curLength : maxLength);
                        int colounwidth = 256 * maxLength;
                        sheet.SetColumnWidth(j, colounwidth);
                        try
                        {
                            cell = row.CreateCell(j);//创建第i行的第j列
                            #region 插入第j列的数据
                            try
                            {
                                NpoiDataType dtype = datatypes[j];
                                switch (dtype)
                                {
                                    case NpoiDataType.String:
                                        {
                                            cell.SetCellType(CellType.String);
                                            cell.SetCellValue(columnValue.ToString());
                                        }
                                        break;
                                    case NpoiDataType.Datetime:
                                        {
                                            cell.SetCellType(CellType.String);
                                            cell.SetCellValue(columnValue.ToString());
                                        }
                                        break;
                                    case NpoiDataType.Numeric:
                                        {
                                            cell.SetCellType(CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(columnValue));
                                        }
                                        break;
                                    case NpoiDataType.Bool:
                                        {
                                            cell.SetCellType(CellType.Boolean);
                                            cell.SetCellValue(Convert.ToBoolean(columnValue));
                                        }
                                        break;
                                    case NpoiDataType.Richtext:
                                        {
                                            cell.SetCellType(CellType.Formula);
                                            cell.SetCellValue(columnValue.ToString());
                                        }
                                        break;
                                }
                            }
                            catch
                            {
                                cell.SetCellType(CellType.String);
                                cell.SetCellValue(columnValue.ToString());
                            }
                            #endregion

                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
                #endregion

                //using (FileStream fs = new FileStream(@SaveFileName, FileMode.OpenOrCreate))//生成文件在服务器上
                //{
                //    wb.Write(fs);
                //}
                //string SaveFileName = "output.xls";
                using (FileStream fs = new FileStream(@SaveFileName, FileMode.OpenOrCreate, FileAccess.Write))//生成文件在服务器上
                {
                    wb.Write(fs);
                    Console.WriteLine("文件保存成功！" + SaveFileName);
                }

                return true;
            }
            catch (Exception)
            {
                Console.WriteLine("文件保存成功！" + SaveFileName);
                return false;
            }

        }
        #endregion

        #region 从DataSet导出到MemoryStream流2007
        /// <summary>
               /// 从DataSet导出到MemoryStream流2007
               /// </summary>
               /// <param name="SaveFileName">文件保存路径</param>
               /// <param name="SheetName">Excel文件中的Sheet名称</param>
               /// <param name="ds">存储数据的DataSet</param>
               /// <param name="startRow">从哪一行开始写入，从0开始</param>
               /// <param name="datatypes">DataSet中的各列对应的数据类型</param>
        public bool CreateExcel2007(string SaveFileName, string SheetName, DataSet ds, int startRow, params NpoiDataType[] datatypes)
        {
            try
            {
                if (startRow < 0)
                {
                    startRow = 0;
                }

                XSSFWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet(SheetName);
                //sheet.SetColumnWidth(0, 50 * 256);
                //sheet.SetColumnWidth(1, 100 * 256);
                IRow row;
                ICell cell;
                DataRow dr;
                int j;
                int maxLength = 0;
                int curLength = 0;
                object columnValue;
                DataTable dt = ds.Tables[0];
                if (datatypes.Length < dt.Columns.Count)
                {
                    datatypes = new NpoiDataType[dt.Columns.Count];
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        string dtcolumntype = dt.Columns[i].DataType.Name.ToLower();
                        switch (dtcolumntype)
                        {
                            case "string":
                                datatypes[i] = NpoiDataType.String;
                                break;
                            case "datetime":
                                datatypes[i] = NpoiDataType.Datetime;
                                break;
                            case "boolean":
                                datatypes[i] = NpoiDataType.Bool;
                                break;
                            case "double":
                                datatypes[i] = NpoiDataType.Numeric;
                                break;
                            default:
                                datatypes[i] = NpoiDataType.String;
                                break;
                        }
                    }
                }

                #region 创建表头
                row = sheet.CreateRow(0);//创建第i行
                ICellStyle style1 = wb.CreateCellStyle();//样式
                IFont font1 = wb.CreateFont();//字体

                font1.Color = HSSFColor.White.Index;//字体颜色
                font1.Boldweight = (short)FontBoldWeight.Bold;//字体加粗样式
                                                              //style1.FillBackgroundColor = HSSFColor.WHITE.index;//GetXLColour(wb, LevelOneColor);// 设置图案色
                style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;//GetXLColour(wb, LevelOneColor);// 设置背景色
                style1.FillPattern = FillPattern.SolidForeground;
                style1.SetFont(font1);//样式里的字体设置具体的字体样式
                style1.Alignment = HorizontalAlignment.Center;//文字水平对齐方式
                style1.VerticalAlignment = VerticalAlignment.Center;//文字垂直对齐方式
                row.HeightInPoints = 25;
                for (j = 0; j < dt.Columns.Count; j++)
                {
                    columnValue = dt.Columns[j].ColumnName;
                    curLength = Encoding.Default.GetByteCount(columnValue.ToString());
                    maxLength = (maxLength < curLength ? curLength : maxLength);
                    int colounwidth = 256 * maxLength;
                    sheet.SetColumnWidth(j, colounwidth);
                    try
                    {
                        cell = row.CreateCell(j);//创建第0行的第j列
                        cell.CellStyle = style1;//单元格式设置样式

                        try
                        {
                            //cell.SetCellType(CellType.STRING);
                            cell.SetCellValue(columnValue.ToString());
                        }
                        catch { }

                    }
                    catch
                    {
                        continue;
                    }
                }
                #endregion

                #region 创建每一行
                for (int i = startRow; i < ds.Tables[0].Rows.Count; i++)
                {
                    dr = ds.Tables[0].Rows[i];
                    row = sheet.CreateRow(i + 1);//创建第i行
                    for (j = 0; j < dt.Columns.Count; j++)
                    {
                        columnValue = dr[j];
                        curLength = Encoding.Default.GetByteCount(columnValue.ToString());
                        maxLength = (maxLength < curLength ? curLength : maxLength);
                        int colounwidth = 256 * maxLength;
                        sheet.SetColumnWidth(j, colounwidth);
                        try
                        {
                            cell = row.CreateCell(j);//创建第i行的第j列
                            #region 插入第j列的数据
                            try
                            {
                                NpoiDataType dtype = datatypes[j];
                                switch (dtype)
                                {
                                    case NpoiDataType.String:
                                        {
                                            //cell.SetCellType(CellType.STRING);
                                            cell.SetCellValue(columnValue.ToString());
                                        }
                                        break;
                                    case NpoiDataType.Datetime:
                                        {
                                            // cell.SetCellType(CellType.STRING);
                                            cell.SetCellValue(columnValue.ToString());
                                        }
                                        break;
                                    case NpoiDataType.Numeric:
                                        {
                                            //cell.SetCellType(CellType.NUMERIC);
                                            cell.SetCellValue(Convert.ToDouble(columnValue));
                                        }
                                        break;
                                    case NpoiDataType.Bool:
                                        {
                                            //cell.SetCellType(CellType.BOOLEAN);
                                            cell.SetCellValue(Convert.ToBoolean(columnValue));
                                        }
                                        break;
                                    case NpoiDataType.Richtext:
                                        {
                                            // cell.SetCellType(CellType.FORMULA);
                                            cell.SetCellValue(columnValue.ToString());
                                        }
                                        break;
                                }
                            }
                            catch
                            {
                                //cell.SetCellType(HSSFCell.CELL_TYPE_STRING);
                                cell.SetCellValue(columnValue.ToString());
                            }
                            #endregion

                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
                #endregion

                //using (FileStream fs = new FileStream(@SaveFileName, FileMode.OpenOrCreate))//生成文件在服务器上
                //{
                //    wb.Write(fs);
                //}
                //string SaveFileName = "output.xlsx";
                using (FileStream fs = new FileStream(SaveFileName, FileMode.OpenOrCreate, FileAccess.Write))//生成文件在服务器上
                {
                    wb.Write(fs);
                    Console.WriteLine("文件保存成功！" + SaveFileName);
                }
                return true;
            }
            catch (Exception)
            {
                Console.WriteLine("文件保存失败！" + SaveFileName);
                return false;
            }

        }
        #endregion

        private short GetXLColour(HSSFWorkbook workbook, System.Drawing.Color SystemColour)
        {
            short s = 0;
            HSSFPalette XlPalette = workbook.GetCustomPalette();
            NPOI.HSSF.Util.HSSFColor XlColour = XlPalette.FindColor(SystemColour.R, SystemColour.G, SystemColour.B);
            if (XlColour == null)
            {
                if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 255)
                {
                    if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 64)
                    {
                        //NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE= 64;
                        //NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE += 1;
                        XlColour = XlPalette.AddColor(SystemColour.R, SystemColour.G, SystemColour.B);
                    }
                    else
                    {
                        XlColour = XlPalette.FindSimilarColor(SystemColour.R, SystemColour.G, SystemColour.B);
                    }
                    s = XlColour.Indexed;
                }
            }
            else
            {
                s = XlColour.Indexed;
            }

            return s;
        }

        #region 读Excel-根据NpoiDataType创建的DataTable列的数据类型
        /// <summary>
               /// 读Excel-根据NpoiDataType创建的DataTable列的数据类型
               /// </summary>
               /// <param name="datatype"></param>
               /// <returns></returns>
        private Type GetDataTableType(NpoiDataType datatype)
        {
            Type tp = typeof(string);//Type.GetType("System.String")
            switch (datatype)
            {
                case NpoiDataType.Bool:
                    tp = typeof(bool);
                    break;
                case NpoiDataType.Datetime:
                    tp = typeof(DateTime);
                    break;
                case NpoiDataType.Numeric:
                    tp = typeof(double);
                    break;
                case NpoiDataType.Error:
                    tp = typeof(string);
                    break;
                case NpoiDataType.Blank:
                    tp = typeof(string);
                    break;
            }
            return tp;
        }
        #endregion

        #region 读Excel-得到不同数据类型单元格的数据
        /// <summary>
               /// 读Excel-得到不同数据类型单元格的数据
               /// </summary>
               /// <param name="datatype">数据类型</param>
               /// <param name="row">数据中的一行</param>
               /// <param name="column">哪列</param>
               /// <returns></returns>
        private object GetCellData(NpoiDataType datatype, IRow row, int column)
        {

            switch (datatype)
            {
                case NpoiDataType.String:
                    try
                    {
                        return row.GetCell(column).DateCellValue;
                    }
                    catch
                    {
                        try
                        {
                            return row.GetCell(column).StringCellValue;
                        }
                        catch
                        {
                            return row.GetCell(column).NumericCellValue;
                        }
                    }
                case NpoiDataType.Bool:
                    try { return row.GetCell(column).BooleanCellValue; }
                    catch { return row.GetCell(column).StringCellValue; }
                case NpoiDataType.Datetime:
                    try { return row.GetCell(column).DateCellValue; }
                    catch { return row.GetCell(column).StringCellValue; }
                case NpoiDataType.Numeric:
                    try { return row.GetCell(column).NumericCellValue; }
                    catch { return row.GetCell(column).StringCellValue; }
                case NpoiDataType.Richtext:
                    try { return row.GetCell(column).RichStringCellValue; }
                    catch { return row.GetCell(column).StringCellValue; }
                case NpoiDataType.Error:
                    try { return row.GetCell(column).ErrorCellValue; }
                    catch { return row.GetCell(column).StringCellValue; }
                case NpoiDataType.Blank:
                    try { return row.GetCell(column).StringCellValue; }
                    catch { return ""; }
                default: return "";
            }
        }
        #endregion

        #region 获取单元格数据类型
        /// <summary>
               /// 获取单元格数据类型
               /// </summary>
               /// <param name="hs"></param>
               /// <returns></returns>
        private NpoiDataType GetCellDataType(ICell hs)
        {
            NpoiDataType dtype;
            DateTime t1;
            string cellvalue = "";

            switch (hs.CellType)
            {
                case CellType.Blank:
                    dtype = NpoiDataType.String;
                    cellvalue = hs.StringCellValue;
                    break;
                case CellType.Boolean:
                case CellType.Numeric:
                    dtype = NpoiDataType.Numeric;
                    cellvalue = hs.NumericCellValue.ToString();
                    break;
                case CellType.String:
                    dtype = NpoiDataType.String;
                    cellvalue = hs.StringCellValue;
                    break;
                case CellType.Error:
                    dtype = NpoiDataType.Error;
                    break;
                case CellType.Formula:
                default:
                    dtype = NpoiDataType.Datetime;
                    break;
            }
            if (cellvalue != "" && DateTime.TryParse(cellvalue, out t1))
            {
                dtype = NpoiDataType.Datetime;
            }

            return dtype;
        }
        #endregion



        #region 测试代码


        #endregion
    }

    #region 枚举(Excel单元格数据类型)
    /// <summary>
       /// 枚举(Excel单元格数据类型)
       /// </summary>
    public enum NpoiDataType
    {
        /// <summary>
               /// 字符串类型-值为1
               /// </summary>
        String,
        /// <summary>
               /// 布尔类型-值为2
               /// </summary>
        Bool,
        /// <summary>
               /// 时间类型-值为3
               /// </summary>
        Datetime,
        /// <summary>
               /// 数字类型-值为4
               /// </summary>
        Numeric,
        /// <summary>
               /// 复杂文本类型-值为5
               /// </summary>
        Richtext,
        /// <summary>
               /// 空白
               /// </summary>
        Blank,
        /// <summary>
               /// 错误
               /// </summary>
        Error
    }
    #endregion
}
