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

            DataTable dtExcel = new DataTable();
            string conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @"Lista studenti.xlsx" + "; Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  

            OleDbConnection con = new OleDbConnection(conn);

            OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Foaie1$]", con); //here we read data from sheet1  
            oleAdpt.Fill(dtExcel); //fill excel data into dataTable  

            dataGridView1.Visible = true;
            dataGridView1.DataSource = dtExcel;

        }
}
}
