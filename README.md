A simple unfinished program that transfers data from the Oracle database to the Excel table.

____________________________________________________________________________________________

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
        OraConnect.Close();
        OraConnect.Dispose();
    }
}           

____________________________________________________________________________________________       
    
