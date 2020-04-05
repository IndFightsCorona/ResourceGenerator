using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Reflection;

namespace ExcelToResx
{
    class ExcelAccess
    {
        private string SourceFile { get; }
        private OleDbConnection OledbConn { get; set; }
        public OleDbCommand OledbCmd { get; private set; }

        public ExcelAccess()
        {
            SourceFile = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SourceFile", @"FightCorona_ResourceValues.xlsx");
            InitializeOledbConnection();
        }
        private void InitializeOledbConnection()
        {
            string connString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""", SourceFile);
            OledbConn = new OleDbConnection(connString);
        }
        public DataSet GetExcelWorkSheetsName()
        {
            OledbConn.Open();
            DataTable mySheets = OledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            OledbConn.Close();
            DataSet ds = new DataSet();
            DataTable dt;

            for (int i = 0; i < mySheets.Rows.Count; i++)
            {
                dt = GetExcelDataAsTable(mySheets.Rows[i]["TABLE_NAME"].ToString());
                ds.Tables.Add(dt);
            }
            return ds;
        }

        public DataTable GetExcelDataAsTable(string WorkSheetName)
        {
            try
            {

                DataTable dt = new DataTable();
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = $"Select * from [{WorkSheetName}]";
                    comm.Connection = OledbConn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        dt.TableName = WorkSheetName.Remove(WorkSheetName.Length-1);
                        return dt;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;

            }
        }
    }
}
