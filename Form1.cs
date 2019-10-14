using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace ExcelHelperDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (!string.IsNullOrEmpty(file.FileName))
                {
                    //这里有三个参数，每一个都看一下
                    //var dt = ExcelHelper.ExcelToDataTable(file.FileName);
                    //var dt = ExcelHelper.ExcelToDataTable(file.FileName, "Sheet1");
                    var dt = ExcelHelper.ExcelToDataTable(file.FileName, "Sheet1", false);

                    if (dt != null)
                    {
                        MessageBox.Show("data rows " + dt.Rows.Count);
                        // dt 拿到数据后  自行处理
                        // todo
                    }
                    else
                    {
                        MessageBox.Show("no data");
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var source = GetExportSource();

            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Title = "请选择要保存的文件路径";
            //初始化保存目录，默认exe文件目录
            fileDialog.InitialDirectory = Application.StartupPath;
            //设置保存文件的类型
            fileDialog.Filter = "文本文件|*.xls|文本文件|*.xlsx|文本文件|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                //导出excel代码
                HSSFWorkbook workbook = new HSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("导出一个demo");

                // 创建第一行 -- 列头
                IRow rowTitle = sheet.CreateRow(0);
                ICell cellTitle0 = rowTitle.CreateCell(0);
                cellTitle0.SetCellValue("编号");

                ICell cellTitle1 = rowTitle.CreateCell(1);
                cellTitle1.SetCellValue("内容");

                int rowIndex = 1;
                // 创建行  --- 导入内容
                foreach (DataRow ri in source.Rows)
                {
                    IRow rowData = sheet.CreateRow(rowIndex++);
                    ICell cellData0 = rowData.CreateCell(0);
                    cellData0.SetCellValue(ri[0].ToString());

                    ICell cellData1 = rowData.CreateCell(1);
                    cellData1.SetCellValue(ri[1].ToString());

                }
                var fileName = !fileDialog.FileName.EndsWith(".xls") ? fileDialog.FileName + ".xls" : fileDialog.FileName;
                FileStream stream = new FileStream(fileName, FileMode.Create);
                workbook.Write(stream);
                stream.Close();
                stream.Dispose();
            }
        }

        private DataTable GetExportSource()
        {
            var dt = new DataTable();
            DataColumn column1 = new DataColumn("序号", typeof(string));
            DataColumn column2 = new DataColumn("内容", typeof(string));
            dt.Columns.AddRange(new DataColumn[] { column1, column2 });
            for (var i = 0; i < 10; i++)
            {
                var newRow = dt.NewRow();
                newRow["序号"] = i + 1;
                newRow["内容"] = "内容" + i.ToString();
                dt.Rows.Add(newRow);
            }
            return dt;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //这个地方并没有设置选择文件。节省时间

            //1.读取Execl数据
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (!string.IsNullOrEmpty(file.FileName))
                {
                    var ds = new NpoiHelper().ReadExcel(file.FileName, 1);
                    MessageBox.Show(ds.Tables.Count.ToString());
                }
            }

        }
    }
}
