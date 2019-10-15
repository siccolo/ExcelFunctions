using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

//
using ExcelDataReader;
//

namespace ConsoleApp1
{
    public class Excel
    {
        public DataTable ReadExcelFile(string sheetName, string path)
        {

            using (OleDbConnection conn = new OleDbConnection())
            {
                DataTable dt = new DataTable();
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";

                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }
                }
            }
        }

        public  void CheckExcelColumnName(string sourceFile)
        {
            if (!System.IO.File.Exists(sourceFile))
            {
                throw new System.ArgumentException("source file is missing");
            }

            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic excel = Activator.CreateInstance(excelType);

                if (excel == null)
                {
                    throw new System.Exception("Excel is missing");
                }

                excel.Workbooks.Open(sourceFile);
                try
                {
                    dynamic workbook = excel.Workbooks[1];
                    dynamic worksheets = workbook.Sheets;
                    dynamic worksheet = worksheets[1];
                    dynamic cells = worksheet.Cells;

                    int xlCellTypeLastCell = 11;
                    dynamic lastfilledcell = cells.SpecialCells(xlCellTypeLastCell, Type.Missing);
                    int lastcolumn = lastfilledcell.Column;

                    dynamic range = worksheet.Range("A1", "A1");//
                    var v = range.EntireRow.Value;
                    for(int c=1;c<= lastcolumn;c++)
                    {
                        if ( v[1,c]==null)
                        {
                            Console.WriteLine("excel sheet is missing column name for column #" + c);
                            break;
                        }
                    }

                    workbook.Close();

                    lastfilledcell = null;
                    cells = null;
                    range = null;
                    worksheet = null;
                    worksheets = null;
                    workbook = null;

                }
                catch {; }

                excel.Quit();
                //try { System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); } catch {; }
                try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excel); } catch {; }
                excel = null;

            }
            catch(Exception ex)
            {

            }
        }

        public void ConvertToCSV(string sourceFile, string targetFile)
        {
            using (var stream = System.IO.File.Open(sourceFile, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                //add ExcelDataReader and ExcelDataReader.DataSet
                //Reading from a OpenXml Excel file (2007 format; *.xlsx)
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    //DataSet result = reader.AsDataSet();
                    DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });
                    if (result.Tables.Count > 0)
                    {
                        System.Text.StringBuilder output = new StringBuilder();
                        DataTable table = result.Tables[0];
                        //save column names
                        output.AppendLine(String.Join(",", table.Columns.Cast<System.Data.DataColumn>().ToList()));
                        //save all rows
                        foreach (System.Data.DataRow dr in table.Rows)
                        {
                            output.AppendLine(String.Join(",", dr.ItemArray.Select(f=>f.ToString() ).ToList()   ) );
                        }
                        System.IO.File.WriteAllText(targetFile, output.ToString());
                    }
                }
            }
        }
    }
}
