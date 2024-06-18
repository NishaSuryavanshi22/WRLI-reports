using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;



namespace CSCUtils
{
    class Utils
    {
        DataSet dsPolicy = new DataSet();
        DataTable datatab = new DataTable(); // Create a new Data table
        public static string GeteDeliveryDB()
        {
            string sDBName = ConfigurationManager.AppSettings["dbConnectionString"];
            return sDBName;
        }

        public static string GetConnectionString() 
        {
            try
            {
                string sDBName = ConfigurationManager.AppSettings["dbConnectionString"];
                string sDBServer = ConfigurationManager.AppSettings["Data Source"];
                string sDBConnectionString = ConfigurationManager.AppSettings["dbConnectionString"].ToString().Replace("%User ID%", sDBName);
                sDBConnectionString = sDBConnectionString.Replace("%Data Source%", sDBServer);
                return sDBConnectionString;
            }
            catch
            {
                return string.Empty;
            }

        }

        public static T ConvertFromDBVal<T>(object obj)
        {
            if (obj == null || obj == DBNull.Value)
            {
                return default(T); // returns the default value for the type
            }
            else
            {
                return (T)obj;
            }
        }
        
        public static string NullToBlank(string str)
        {
            string result = "";
            if (!String.IsNullOrEmpty(str))
                result = str;
            return result.Trim();
        }

        public static string ConvertStringToHex(string asciiString)
        {
            string hex = "";
            try
            {
                foreach (char c in asciiString)
                {
                    int tmp = c;
                    hex += String.Format("{0:x2}", (uint)System.Convert.ToUInt32(tmp.ToString()));
                }
            }
            catch { }
            return hex;
        }

        public static string ConvertHexToString(string HexValue)
        {
            string StrValue = "";
            try
            {
                while (HexValue.Length > 0)
                {
                    StrValue += System.Convert.ToChar(System.Convert.ToUInt32(HexValue.Substring(0, 2), 16)).ToString();
                    HexValue = HexValue.Substring(2, HexValue.Length - 2);
                }
            }
            catch { }
            return StrValue;
        }

        public static byte[] ExportToCSVFileOpenXML(DataTable dt)
        {
            DataSet ds = new DataSet();
            DataTable dtCopy = new DataTable();
            dtCopy = dt.Copy();
            ds.Tables.Add(dtCopy);
            try
            {
                byte[] returnBytes = null;
                MemoryStream mem = new MemoryStream();
                var workbook = SpreadsheetDocument.Create(mem, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
                {
                    var workbookPart = workbook.AddWorkbookPart();
                    workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                    workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();
                    foreach (System.Data.DataTable table in ds.Tables)
                    {
                        var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                        var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                        sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                        DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                        string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                        uint sheetId = 1;
                        if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                        {
                            sheetId =
                                sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                        }

                        DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                        sheets.Append(sheet);
                        

                        DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                        List<String> columns = new List<string>();
                        foreach (System.Data.DataColumn column in table.Columns)
                        {
                            columns.Add(column.ColumnName);

                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                            
                            headerRow.AppendChild(cell);
                        }


                        sheetData.AppendChild(headerRow);

                        foreach (System.Data.DataRow dsrow in table.Rows)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                            foreach (String col in columns)
                            {
                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                                newRow.AppendChild(cell);
                            }

                            sheetData.AppendChild(newRow);
                        }

                    }
                }
                
                workbook.WorkbookPart.Workbook.Save();
                workbook.Close();

                returnBytes = mem.ToArray();

                return returnBytes;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static byte[] ExportToCSVFileOpenXML_2(DataTable dt)
        {
            byte[] returnBytes = null;
            using (MemoryStream mem = new MemoryStream())
            {
                var workbook = SpreadsheetDocument.Create(mem, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);

                // your code

                workbook.WorkbookPart.Workbook.Save();
                workbook.Close();

                returnBytes = mem.ToArray();
            }

            return returnBytes;
        }

        public static string DateToLPDate(string sDate)
        {
            string sReturn = "19000101";
            try
            {

                string[] fromDate = sDate.Split('/');
                sReturn = fromDate[2] + fromDate[0].PadLeft(2, '0') + fromDate[1].PadLeft(2, '0');
            }
            catch { };
            return sReturn;
        }

        
        
    }
}
