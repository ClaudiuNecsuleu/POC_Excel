using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace POC_Excel
{
    using Excel = Microsoft.Office.Interop.Excel;
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Workbook workbook = excel.Workbooks.Open(@"C:\Users\Toni\Desktop\POC_Excel\POC_Excel\bin\Debug\Lista studenti.xlsx", ReadOnly: false, Editable: true);
            Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;

            for (int i = 2; i <= 4; i++)
            {
                
                Range t = worksheet.Cells[i,2];
                t.Value = ((string)t.Value).ToUpper();
            }

            excel.Application.ActiveWorkbook.Save();
            excel.Application.Quit();
            excel.Quit();

        }
}
}
