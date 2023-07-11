using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Odbc;
using Oracle.ManagedDataAccess.Client;

namespace OracleConnector
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SelectfromDB Named = new SelectfromDB(textBox1.Text);

            string sql = "SELECT OBJECT_ID, CLASS_ID, CN_CODE, CN_NAME, CN_MASS FROM SMTEST.TN_ITEMS WHERE CN_CODE like '%";
            DataSet ds = new DataSet();
            Named.DB(ds, sql);
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            this.Close();
            Dispose();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }
    }
}
