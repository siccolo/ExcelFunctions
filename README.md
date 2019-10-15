# ExcelFunctions

Provider=Microsoft.ACE.OLEDB.12.0; -> OleDbCommand -> OleDbDataAdapter

Excel.Application -> Open() excel file

ExcelDataReader and ExcelDataReader.DataSet -> ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader() -> reader.AsDataSet(new ExcelDataSetConfiguration())
