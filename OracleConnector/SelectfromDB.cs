using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.Odbc;
using Oracle.ManagedDataAccess.Client;
using System.Windows.Forms;


namespace OracleConnector
{
    struct Selected
    {
        public int OBJECT_ID { get; set; }
        public int CLASS_ID { get; set; }
        public string CN_CODE { get; set; }
        public string CN_NAME { get; set; }
    }
    class SelectfromDB
    {
        public static string CN_CODE;
        public static OracleConnection OraConnect;
        public static OracleCommand OraCommand;
        public SelectfromDB(string cn_cd)
        {
            CN_CODE = cn_cd;
        }
        public void ExcelImporting()
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            oXL = new Excel.Application();
            oWB = oXL.Workbooks.Add(Type.Missing);
            oSheet = (Excel.Worksheet)oXL.Sheets[1];
            oSheet.Name = "Selected";
            foreach (DataTable table in ds.Tables)
            {
                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    oSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        oSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }
            foreach (Excel.Worksheet wrkst in oWB.Worksheets)
            {
                Excel.Range usedrange = wrkst.UsedRange;
                usedrange.Columns.AutoFit();
                oSheet.get_Range("A").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }
            oXL.Visible = true;
            oXL.UserControl = true;
        }
         public void DB(DataSet ds, string select1)
        {
            string str = "(DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = )(PORT = )))(CONNECT_DATA = (SERVICE_NAME = )(SERVER = DEDICATED)))";
            string connString = "Data Source=" + str + ";User ID=;Password=;";
            OraConnect = new OracleConnection(connString);
            string select = select1 + CN_CODE + "%'";
            try
            {
                OraConnect.Open();
                if (OraConnect.State == ConnectionState.Open)
                {
                    OracleCommand cmd = OraConnect.CreateCommand();
                    cmd.CommandText = select;
                    OracleDataAdapter da = new OracleDataAdapter(cmd);
                    ds.Tables.Add();
                    da.Fill(ds);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (OraConnect != null)
                {
                    if (OraConnect.State == ConnectionState.Open)
                    {
                        try
                        {
                            Excel.Application oXL;
                            Excel._Workbook oWB;
                            Excel._Worksheet oSheet;
                            oXL = new Excel.Application();
                            oWB = oXL.Workbooks.Add(Type.Missing);
                            oSheet = (Excel.Worksheet)oXL.Sheets[1];
                            oSheet.Name = "Selected";
                            foreach (DataTable table in ds.Tables)
                            {
                                for (int i = 1; i < table.Columns.Count + 1; i++)
                                {
                                    oSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                                }
                                for (int j = 0; j < table.Rows.Count; j++)
                                {
                                    for (int k = 0; k < table.Columns.Count; k++)
                                    {
                                        oSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                                    }
                                }
                            }
                            foreach (Excel.Worksheet wrkst in oWB.Worksheets)
                            {
                                Excel.Range usedrange = wrkst.UsedRange;
                                usedrange.Columns.AutoFit();
                                oSheet.get_Range("A").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            }
                            oXL.Visible = true;
                            oXL.UserControl = true;
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show(e.Message);
                        }
                        OraConnect.Close();
                        OraConnect.Dispose();
                    }
                }
            }
        }
    }
}