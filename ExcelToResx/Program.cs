using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Resources;
using System.Text.RegularExpressions;

namespace ExcelToResx
{
    class Program
    {
        static void Main(string[] args)
        {
           
            ExcelAccess exclInstance = new ExcelAccess();
            DataSet dataSet = exclInstance.GetExcelWorkSheetsName();
            string oututFolder = ConfigurationManager.AppSettings["OutputFolder"];
            if (Directory.Exists(oututFolder))
                Directory.Delete(oututFolder, true);
            GenerateResxFile(dataSet, oututFolder);
        }
        private static void GenerateResxFile(DataSet dataSet, string oututFolder)
        {
            foreach (DataTable dtEcelData in dataSet.Tables)
            {
                if (!Directory.Exists(Path.Combine(oututFolder, String.Format($"{dtEcelData.TableName}"))))
                    Directory.CreateDirectory(Path.Combine(oututFolder, String.Format($"{dtEcelData.TableName}")));
                foreach (DataColumn excelColumn in dtEcelData.Columns)
                {

                    if (excelColumn.ColumnName != "Key" && excelColumn.ColumnName != "Default Value" && !Regex.IsMatch(excelColumn.ColumnName, "^[A-Za-z]{1}[0-9]{1}"))
                    {
                        string fileName = Path.Combine(oututFolder, String.Format($"{dtEcelData.TableName}/{dtEcelData.TableName}.{excelColumn.ColumnName.ToLower()}.resx"));

                        using (ResXResourceWriter resx = new ResXResourceWriter(fileName))
                        {
                            foreach (DataRow dRow in dtEcelData.Rows)
                            {
                                if (!string.IsNullOrEmpty(dRow.Field<string>("Key")) && !string.IsNullOrEmpty(dRow.Field<string>(excelColumn.ColumnName)))
                                    resx.AddResource(dRow.Field<string>("Key"), dRow.Field<string>(excelColumn.ColumnName));
                            }
                        }
                    }
                }
            }
        }
    }
}
