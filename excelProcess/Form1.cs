using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelOperation;

namespace excelProcess
{
    public partial class Form1 : Form
    {
        private string ExcelConnStr;
        private string ExcelStr;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fName;
            openFileDialog1.InitialDirectory = "E:\\考勤";//注意这里写路径时要用c:\\而不是c:\
            openFileDialog1.Filter = "Excel2003文件|*.xls|Excel文件|*.xlsx";
            openFileDialog1.RestoreDirectory = false;
            openFileDialog1.FilterIndex = 1;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fName = openFileDialog1.FileName;//全路径
                getDataTable(fName);


            }
        }

        private DataTable getDataTable(string path)
        {
            ExcelConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1';";
            ExcelStr = " SELECT * FROM [Sheet 1$] ";
            DataTable dt = ExcelOperation.Operation.GetExcelSearchTable(ExcelConnStr, ExcelStr);
            return dt;
        }
    }
}
