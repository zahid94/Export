using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;

namespace Export.Models
{
    public static class Utility
    {
        public static DataSet ConvertXSLXtoDataSet(string strFilePath, string connString)
        {
            OleDbConnection oledbConn = new OleDbConnection(connString);
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            try
            {

                oledbConn.Open();
                using (DataTable Sheets = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null))
                {

                    for (int i = 0; i < Sheets.Rows.Count; i++)
                    {
                        string worksheets = Sheets.Rows[i]["TABLE_NAME"].ToString();
                        OleDbCommand cmd = new OleDbCommand(String.Format("SELECT * FROM [{0}]", worksheets), oledbConn);
                        OleDbDataAdapter oleda = new OleDbDataAdapter();
                        oleda.SelectCommand = cmd;
                        DataTable t = new DataTable(worksheets);

                        oleda.Fill(t);
                        ds.Tables.Add(t);
                    }


                }

            }
            catch (Exception)
            {
            }
            finally
            {

                oledbConn.Close();
            }

            return ds;

        }
    }
}