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
    public partial class CurrentDateReport_Averages : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");

       // SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        string agent = "WRE";
        DataSet dsPolicy = new DataSet();
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
         string[] strRowValue = new string[] { };
         DataTable datatab = new DataTable(); // Create a new Data table

         
         public static DataTable dataPolicy = new DataTable();
         
        
         bool Yearly = false;
        

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
            tblgrid.Visible = false;
                //txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                //txtFromdate.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                //txtFromdate.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                else
                    sCompany = "15";
                if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                    AgentID = Session["LoginID"].ToString();

                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();
                Yearly = false;
                idGroup.Value = "false";
                if (Request.QueryString["GroupYTD"] != null && Request.QueryString["GroupYTD"].ToString() != "")
                {
                    Yearly = Convert.ToBoolean(Request.QueryString["GroupYTD"].ToString());
                    idGroup.Value = (Request.QueryString["GroupYTD"].ToString());
                }
            if (Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "")
                FromDate = Request.QueryString["fromdate"].ToString();
            //if (Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "")
            //    ToDate = Request.QueryString["todate"].ToString();
            //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Initial Catalog=WRE_AGENT");


            InvokeSP();  
        }

        protected void InvokeSP()
        {

            //string[] fromDate = txtFromdate.Text.Split('/');
            //FromDate = fromDate[2] + fromDate[0] + fromDate[1];

            //string[] toDate = txtTo.Text.Split('/');
            //ToDate = toDate[2] + toDate[0] + toDate[1];

            con.Open();

            SqlCommand commPolicy = new SqlCommand();

            commPolicy.Connection = con;
            //commPolicy.CommandType = CommandType.StoredProcedure;
            //Yearly = true;

            if (idGroup.Value =="true")
            {
                commPolicy.CommandText = "select  '" + FromDate + "'+substring(cast(isnull(ppp.app_received_date,po.app_received_date)as varchar(8)),5,2) as Dispmonth, " +
             " (select count(*) from pending_policy p left outer join policies2 pp on p.policy_number = pp.policy_number and p.company_code = pp.company_code "+
             " where isnull(p.app_received_date,pp.app_received_date) between '" + FromDate + "' and  '" + FromDate + "' " +
             "+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '31' and p.agent_number like 'C%' )   as Starts_With_C_YTD,"+
  
             "(select count(*) from pending_policy p left outer join policies2 pp on p.policy_number = pp.policy_number and p.company_code = pp.company_code "+
             " where isnull(p.app_received_date,pp.app_received_date) between '" + FromDate + "' and  '" + FromDate + "' " +
             " +substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '31' and p.agent_number like 'A%' )   as Starts_With_A_YTD "+
 
             "from pending_policy ppp left outer join policies2 po on ppp.policy_number = po.policy_number and ppp.company_code = po.company_code"+
              " where isnull(ppp.app_received_date,po.app_received_date) is not null "+
              " group by  '" + FromDate + "'+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2) " +
              " order by  '" + FromDate + "'+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2) ";
  
            }
            else
            {
    commPolicy.CommandText = "select '" + FromDate + "'+substring(cast(isnull(ppp.app_received_date,po.app_received_date)as varchar(8)),5,2) as Dispmonth," + 
                    
 " (select count(*) from pending_policy p left outer join policies2 pp on p.policy_number = pp.policy_number and p.company_code = pp.company_code "+
 " where isnull(p.app_received_date,pp.app_received_date) between '" + FromDate + "'+ " +
 " substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '01' and '" + FromDate + "'+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '31' and p.agent_number like 'C%') as Starts_With_C_MTD," +
 
 " (select count(*) from pending_policy p left outer join policies2 pp on p.policy_number = pp.policy_number and p.company_code = pp.company_code  "+
 " where isnull(p.app_received_date,pp.app_received_date) between '" + FromDate + "'+" +
 " substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '01' and '" + FromDate + "'+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '31' and p.agent_number like 'A%') as Starts_With_A_MTD " +
  
 " from pending_policy ppp left outer join policies2 po on ppp.policy_number = po.policy_number and ppp.company_code = po.company_code "+
  " where isnull(ppp.app_received_date,po.app_received_date) is not null  "+
  " group by '" + FromDate + "'+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2) "+
  " order by '" + FromDate + "'+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2) ";
  
            }

            /* commPolicy.CommandText += "select cast(year(getdate()) as varchar(4))+substring(cast(isnull(ppp.app_received_date,po.app_received_date)as varchar(8)),5,2) as dispmonth, " +
 " (select count(*) from pending_policy p left outer join policies2 pp on p.policy_number = pp.policy_number and p.company_code = pp.company_code where isnull(p.app_received_date,pp.app_received_date) between" + "'" + fromDate[2] + "' and cast(year(getdate()) as varchar(4))+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '31' and p.agent_number like 'C%')   as YTD, " +
 " (select count(*) from pending_policy p left outer join policies2 pp on p.policy_number = pp.policy_number and p.company_code = pp.company_code  where isnull(p.app_received_date,pp.app_received_date) between" + "(cast(year(getdate()) as varchar(4))+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '01') and (cast(year(getdate()) " +
             " as varchar(4))+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2)+ '31') and p.agent_number like 'C%') as MTD " +
 " from pending_policy ppp left outer join policies2 po on ppp.policy_number = po.policy_number and ppp.company_code = po.company_code where " +
 " isnull(ppp.app_received_date,po.app_received_date) is not null  group by cast(year(getdate()) as varchar(4))+substring(cast(isnull(pp'p.app_received_date,po.app_received_date) as varchar(8)),5,2)" +
             " order by cast(year(getdate()) as varchar(4))+substring(cast(isnull(ppp.app_received_date,po.app_received_date) as varchar(8)),5,2) ";*/

            
             SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);
            
             //commPolicy.Parameters.AddWithValue("@agentid", AgentID);
             //commPolicy.Parameters.AddWithValue("@company", SqlDbType.VarChar).Value = sCompany;
              

            // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // Create a SQL Data Adapter and assign it the cmd value. 
             datatab = new DataTable(); // Create a new Data table
             dataadapter.Fill(dsPolicy);
             if (dsPolicy != null && dsPolicy.Tables[0] != null)
                 datatab = dsPolicy.Tables[0];
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

                tblgrid.Visible = true;
                grDailyAverages.Visible = true;
                 lblcount.Visible = true;
                 AddEditRows();
                 lblcount.Text = "Total Policy Count: " + datatab.Rows.Count.ToString();
             }

             //grdHandling.DataSource = datatab;
             //grdHandling.DataBind();

             // adPolicy.Fill(dsPolicy);

             con.Close();
         }

         protected void Back_Click(object sender, EventArgs e)
         {
             Response.Redirect("DailyAverages.aspx?Group=Yearly");
         }

         protected void Mtd_Click(object sender, EventArgs e)
         {
         }

         //protected void Button1_Click(object sender, EventArgs e)
         //{
         //    tblgrid.Visible = true;
         //    string selectedComp = "ALL";
         //    string selectedAgent = "ALL";
         //    InvokeSP();
            

         //}

         protected void InitGridColumns(int Rowcount)
         {
             arr_NB[0] = new Int32();
             arr_NB[1] = new Int32();
             arr_NB[2] = new Int32();
             //arr_NB[3] = new Int32();
             //arr_NB[4] = new Int32();

         }
         protected void ReadColumnValues(DataRowView DataRowCurrView, ref string[] strRowValue, int nIndex)
         {
            
            
             //arr_NB[0] += Convert.ToInt32(DataRowCurrView["DISPLAYID"]);
             //arr_NB[1] += Convert.ToInt32(DataRowCurrView["DISPLAYNAME"]);
             /*strRowValue[0] += Convert.ToString(DataRowCurrView["dispmonth"]) + "~";
             if (DataRowCurrView["YTD"] != null && DataRowCurrView["YTD"].ToString().Trim() != "")
                 strRowValue[1] += Convert.ToString(DataRowCurrView["YTD"]) + "~";
             if (DataRowCurrView["MTD"] != null && DataRowCurrView["MTD"].ToString().Trim() != "")
                 strRowValue[2] += Convert.ToString(DataRowCurrView["MTD"]); */
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
             strRowValue = new string[nColCt];
             if ("1" == "1")
                 //InitGridColumns(nColCt);

                
             for (int nIndex = 0; nIndex < nRowct; nIndex++)
             {

                 DataRowCurrView = MyDataView1[nIndex];
                 //List<Int32> stringList = new List<Int32>();
                 //ReadColumnValues(DataRowCurrView, ref strRowValue, nIndex);
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
            ExportToExcel();
            //Export("DailyAverages.xls", this.grDailyAverages);
            //ExportGridView("DailyAverages_Report.xls", this.grdSubmitCount);
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
             
                if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                dataPolicy.DefaultView.Sort = "Dispmonth";
                dataPolicy = dataPolicy.DefaultView.ToTable();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"CurrentDateReport_Averages" + DateTime.Now.ToString();
                Response.AddHeader("Content-Disposition", "inline;filename=" + filename.Replace("/", "").Replace(":", "") + ".xlsx");
                //Call  Export function
                //Response.BinaryWrite(ExportToCSVFileOpenXML(datatab));   
                Response.BinaryWrite(Utils.ExportToCSVFileOpenXML(dataPolicy));
                Response.Flush();
                Response.End();
            }
        }
            
        

        public static void Export(string fileName, GridView gv)
        {
            //Earlier code
            //HttpContext.Current.Response.Clear();
            //HttpContext.Current.Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", fileName));
            //HttpContext.Current.Response.ContentType = "application/ms-excel";
            
            //New code
            string FileName = "111";
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
