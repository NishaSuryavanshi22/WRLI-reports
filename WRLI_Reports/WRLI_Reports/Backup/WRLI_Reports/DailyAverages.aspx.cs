using System;
using System.Collections.Generic;
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
using CSCUtils;
//using ClosedXML.Excel;

//using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace WRLI_Reports
{
    public partial class DailyAverages : System.Web.UI.Page
    {
//        SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        string agent = "WRE";
        DataSet dsPolicy = new DataSet();
        public static DataTable dataPolicy = new DataTable();
        string sCompany = "15";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = "PAY_MONTH,PR.REGION_CODE";
        string OrderDir = "ASC";
        bool bNet = false;
        bool bRegion = false;
        int nRowct = 0;
         string sType = "type";
         Int32[] arr_NB = new Int32[] { };
         string[] strRowValue = new string[3];
         DataTable datatab = new DataTable(); // Create a new Data table

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                if (Session["Validated"] != null && Session["Validated"].ToString() != "A")
                {
                    Response.Redirect("Closed.aspx");
                }
            }
            catch
            {
                Response.Redirect("Closed.aspx");
            } 
            if (!IsPostBack)
            {
                tblgrid.Visible = false;
                txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                else
                    sCompany = "15";
                if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                    AgentID = Session["LoginID"].ToString();

                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();

                //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Initial Catalog=WRE_AGENT");
               
            }
        }

        protected void InvokeSP()
        {

            string[] fromDate = txtFrom.Text.Split('/');
            FromDate = fromDate[2] + fromDate[0] + fromDate[1];

            //string[] toDate = txtTo.Text.Split('/');
            //ToDate = toDate[2] + toDate[0] + toDate[1];

            con.Open();

            SqlCommand commPolicy = new SqlCommand();

            commPolicy.Connection = con;
            //commPolicy.CommandType = CommandType.StoredProcedure;
            commPolicy.CommandText = " select '" + fromDate[2] + "'+ substring(cast(isnull(ppp.app_received_date,po.app_received_date)as varchar(8)),5,2) as Displaymonth,  "+
"(select count(*) from pending_policy p left outer join policies2 pp on p.policy_number = pp.policy_number and p.company_code = pp.company_code  "+
" where isnull(p.app_received_date,pp.app_received_date) between ('" + fromDate[2] + "' +substring(cast(isnull(ppp.app_received_date,"+
" po.app_received_date) as varchar(8)),5,2)+ '01') and ('" + fromDate[2] + "' +substring(cast(isnull(ppp.app_received_date,po.app_received_date) "+
 " as varchar(8)),5,2)+ '31') and p.agent_number like 'C%' ) as 'Agent No. Starts with C', " +

  " (select count(*) from pending_policy p left outer join policies2 pp on p.policy_number = pp.policy_number and p.company_code = pp.company_code where " +
            " isnull(p.app_received_date,pp.app_received_date) between ('" + fromDate[2] + "' +substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '01') " +
   " and ('" + fromDate[2] + "' +substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '31') and p.agent_number like 'A%' ) as 'Agent No. Starts with A' " +
   " from pending_policy ppp left outer join policies2 po on " +
    " ppp.policy_number = po.policy_number and ppp.company_code = po.company_code " +
    " where isnull(ppp.app_received_date,po.app_received_date) is not null "+
            " group by '" + fromDate[2] + "' + substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2) " +
    " order by '" + fromDate[2] + "' + substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)";

            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);
            
            //commPolicy.Parameters.AddWithValue("@agentid", AgentID);
            //commPolicy.Parameters.AddWithValue("@company", SqlDbType.VarChar).Value = sCompany;
              
                
            

           // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // Create a SQL Data Adapter and assign it the cmd value. 
            datatab = new DataTable(); // Create a new Data table
            dataadapter.Fill(datatab);
            //Fill the data table for export excel
            if (datatab != null)
                dataPolicy = datatab;
            if (datatab != null && datatab.Rows.Count == 0)
            {
                DataTable dt = new DataTable();
                DataColumn dc = new DataColumn();

                if (dt.Rows.Count == 0)
                {
                    lblcount.Text = "No Records Found for the selected criteria !!!";
                    dvgrid.Visible = false;
                }
            }
            else
            {

                dvgrid.Visible = true;
                lblcount.Visible = true;
                //Added newly to implement Export to Excel functionality
                if (dataPolicy != null && dataPolicy.DefaultView != null)
                    dataPolicy = dataPolicy.DefaultView.ToTable();
                AddEditRows();
                lblcount.Text = "Total Policy Count: " + datatab.Rows.Count.ToString();
            }

            //grdHandling.DataSource = datatab;
            //grdHandling.DataBind();

            // adPolicy.Fill(dsPolicy);

            con.Close();
        }

        protected void Ytd_Click(object sender, EventArgs e)
        {
            Response.Redirect("CurrentDateReport_Averages.aspx?GroupYTD=true");
        }

        protected void Mtd_Click(object sender, EventArgs e)
        {
            Response.Redirect("CurrentDateReport_Averages.aspx?GroupYTD=false");
        }


        protected void Button1_Click(object sender, EventArgs e)
        {
            tblgrid.Visible = true;
            string selectedComp = "ALL";
            string selectedAgent = "ALL";
            InvokeSP();
            

        }

        protected void InitGridColumns(int Rowcount)
        {
            arr_NB[0] = new Int32();
            arr_NB[1] = new Int32();
            arr_NB[2] = new Int32();
            //arr_NB[3] = new Int32();
            //arr_NB[4] = new Int32();

        }
        protected void ReadColumnValues(DataRowView DataRowCurrView, ref Int32[] arr_NB)
        {

            //arr_NB[0] += Convert.ToInt32(DataRowCurrView["DISPLAYID"]);
            //arr_NB[1] += Convert.ToInt32(DataRowCurrView["DISPLAYNAME"]);
            if (DataRowCurrView["Agent No. Starts with C"] != null && DataRowCurrView["Agent No. Starts with C"].ToString().Trim() != "")
                strRowValue[1] += Convert.ToInt32(DataRowCurrView["Agent No. Starts with C"]) + "~";
            if (DataRowCurrView["Agent No. Starts with A"] != null && DataRowCurrView["Agent No. Starts with A"].ToString().Trim() != "")
                strRowValue[2] += Convert.ToInt32(DataRowCurrView["Agent No. Starts with A"]) + "~";
        }
        protected void AddEditRows()
        {


            DataView MyDataView1 = new DataView(datatab);
            DataRowView DataRowCurrView = null;
            MyDataView1.AllowNew = true;

            //MyDataRowView["active"] = 111;
            //MyDataRowView["sub"] = 222;


            nRowct = datatab.Rows.Count;
            int nColCt = datatab.Columns.Count;
            arr_NB = new Int32[nColCt];
            if ("1" == "1")
                InitGridColumns(nColCt);

             
            for (int nIndex = 0; nIndex < nRowct; nIndex++)
            {
                //strRowValue[nIndex] = new String();
                DataRowCurrView = MyDataView1[nIndex];
                //List<Int32> stringList = new List<Int32>();
                ReadColumnValues(DataRowCurrView, ref arr_NB);
                // End of new logic
            } 
            /* Below line Commented by Siva
            DataRowView MyDataRowView = MyDataView1.AddNew();
            
            //DataRow MyDataRowView = MyDataView1.Table.NewRow();
            //MyDataView1.Table.Rows.InsertAt(MyDataRowView, 0); 

            int position = 0;
            int i = 0;
            MyDataView1.AllowEdit = true;
            MyDataRowView.BeginEdit();
            position = i + 1; //Dont want to insert at the row, but after.
            //if (FilterResultsType == "1")
            if("1" == "1")
            {
                MyDataRowView["DISPLAYID"] = "Total";
                //MyDataRowView["SUB_COUNT"] = arr_NB[3];
                //MyDataRowView["SUB_PREM"] = arr_NB[4];
                MyDataRowView["SUB_PREM"] = "1";
                MyDataRowView["SUB_COUNT"] = "2";
                
            } */
            //grDailyAverages.DataSource = MyDataView1;
                grDailyAverages.DataSource = datatab;
            grDailyAverages.DataBind();

        }

        protected void ddlregion_PreRender(object sender, EventArgs e)
        {
            
        }


        protected void ddlstate_PreRender(object sender, EventArgs e)
        {
            
        }


        protected void ddlpolicydesc_PreRender(object sender, EventArgs e)
        {
            
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            
            //Export the grid data to excel without hitting the Database
            ExportToExcel();
            //Export("DailyAverages_Report.xls", this.grDailyAverages);
        }


        private byte[] ExportToCSVFileOpenXML(DataTable dt)
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
        public byte[] ExportToCSVFileOpenXML_2(DataTable dt)
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

        protected void ExportToExcel()
        {
            //InvokeSP();
            //            dataPolicy = Session["dataPolicy"] as DataTable;

            //dataPolicy = GridView1.DataSource as DataTable;
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                dataPolicy.DefaultView.Sort = "Displaymonth";
                dataPolicy = dataPolicy.DefaultView.ToTable();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"Policy_Report_" + DateTime.Now.ToString("yyyyMMdd");
                Response.AddHeader("Content-Disposition", "inline;filename=" + filename.Replace("/", "").Replace(":", "") + ".xlsx");
                //Call  Export function
                Response.BinaryWrite(Utils.ExportToCSVFileOpenXML(dataPolicy));

                Response.Flush();
                Response.End();
            }

        }

        

        public static void Export(string FileName, GridView gv)
        {
            //Earlier code
            //HttpContext.Current.Response.Clear();
            //HttpContext.Current.Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", fileName));
            //HttpContext.Current.Response.ContentType = "application/ms-excel";
            
            //New code
            //string FileName = "111";
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.ClearContent();
            //HttpContext.Current.Response.Charset = Encoding.UTF8.WebName;
            HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=" + FileName + ".xls");
            HttpContext.Current.Response.AddHeader("Content-Type", "application/Excel");
            //HttpContext.Current.Response.ContentType = "application/octet-stream";
            HttpContext.Current.Response.ContentType = "application/vnd.xlsx";

            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    //  Create a form to contain the grid
                    Table table = new Table();

                    //  add the header row to the table
                    if (gv.HeaderRow != null)
                    {
                     
                        table.Rows.Add(gv.HeaderRow);
                    }

                    //  add each of the data rows to the table
                    foreach (GridViewRow row in gv.Rows)
                    {
                       // PrepareControlForExport(row);
                        table.Rows.Add(row);
                    }

                    //  add the footer row to the table
                    if (gv.FooterRow != null)
                    {
                       // PrepareControlForExport(gv.FooterRow);
                        table.Rows.Add(gv.FooterRow);
                    }

                    //  render the table into the htmlwriter
                    table.RenderControl(htw);

                    //  render the htmlwriter into the response
                    HttpContext.Current.Response.Write(sw.ToString());
                    HttpContext.Current.Response.End();
                }
            }

        }
       
    }

    
       

}
