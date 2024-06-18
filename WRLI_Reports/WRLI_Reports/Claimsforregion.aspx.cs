using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Globalization;
using System.Drawing;
using CSCUtils;
using System.IO;

namespace WRLI_Reports
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);


        DataSet dsPolicy = new DataSet();
        public static DataTable dataPolicy = new DataTable();
        string sCompany = "15";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = "Agent_number ASC";
        string OrderDir = "ASC";
        bool bNet = false;
        bool bRegion = false;
        int nRowct = 0;
        string sType = "type";
        Int32[] arr_NB = new Int32[] { };
        string[] strRowValue = new string[3];
        DataTable datatab = new DataTable(); // Create a new Data table
        string RegionCodeAll = "ALL";
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

            string sCompany = Request.QueryString["COMPANY_CODE"];
            string sRegion = (string)Request.QueryString["Region_Code"];
            string sFromDate1 = (string)Request.QueryString["Fromdate"];
            string sToDate1 = (string)Request.QueryString["Today"];
            string sRider = (string)Request.QueryString["bType"];

            SqlCommand commPolicy = new SqlCommand();

            commPolicy.Connection = con;
            //commPolicy.CommandType = CommandType.StoredProcedure;

            commPolicy.CommandText = "select	AG.AGENT_NUMBER,  "+
         "AGENT_NAME, AGENT_STATUS,"+
         "ISNULL(dbo.PAID_COUNT_FOR_AGENT (AG.AGENT_NUMBER,'" + sRegion + "','" + AgentID + "', '" + sCompany + "', '" + sFromDate1 + "' , '" + sToDate1 + "'),0) as TOTAL_POLICIES, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')) then PAID_CLAIM else 0 END) as TOTAL_PAID_CLAIMS, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')AND CONTESTABLE = 'C1') then PAID_CLAIM else 0 END) as TOTAL_PAID_CLAIMS_C1, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')AND CONTESTABLE = 'C2') then PAID_CLAIM else 0 END) as TOTAL_PAID_CLAIMS_C2, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')AND CONTESTABLE = 'N') then PAID_CLAIM else 0 END) as TOTAL_PAID_CLAIMS_CN, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')) then RESCINDED_CLAIM else 0 END) as TOTAL_RESCINDED, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')AND CONTESTABLE = 'C1') then RESCINDED_CLAIM else 0 END) as TOTAL_RESCINDED_C1, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')AND CONTESTABLE = 'C2') then RESCINDED_CLAIM else 0 END) as TOTAL_RESCINDED_C2, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')AND CONTESTABLE = 'N') then RESCINDED_CLAIM else 0 END) as TOTAL_RESCINDED_CN, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')) then PENDING_CLAIM else 0 END) as TOTAL_PENDING, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')AND CONTESTABLE = 'C1') then PENDING_CLAIM else 0 END) as TOTAL_PENDING_C1, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')AND CONTESTABLE = 'C2') then PENDING_CLAIM else 0 END) as TOTAL_PENDING_C2, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')AND CONTESTABLE = 'N') then PENDING_CLAIM else 0 END) as TOTAL_PENDING_CN, " +
         "SUM(CASE WHEN ((ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')) then (PAID_CLAIM + RESCINDED_CLAIM + PENDING_CLAIM) else 0 END) as TOTAL_CLAIMS, " +
         "ISNULL(AE,0) as AE " +
         "from AGENTS AG " +
         "LEFT OUTER JOIN  CLAIMS_REPORTING CR ON AG.AGENT_NUMBER = CR.AGENT_NUMBER and AG.COMPANY_CODE = CR.COMPANY_CODE " +
         "LEFT OUTER JOIN AE_BYAGENT AE ON AE.NUMBER = AG.AGENT_NUMBER  " +
         "WHERE (AG.COMPANY_CODE='" + sCompany + "') AND AG.AGENT_NUMBER in  " +
         "     (SELECT AH.AGENT_NUMBER " +
         "	   FROM AGENT_HIERLIST AH WHERE (AH.COMPANY_CODE='" + sCompany + "') AND AH.HIERARCHY_AGENT = '" + AgentID + "') " +
         sRider +
          "	   AND (AG.REGION_CODE = '" + sRegion + "') AND	" +
         "      (PAID_POLICY + PAID_CLAIM + RESCINDED_CLAIM + PENDING_CLAIM >= 0) AND " +
         "      ((PAYMENT_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "') OR " +
         "       (ISSUE_DATE BETWEEN '" + sFromDate1 + "' AND '" + sToDate1 + "')) " +
         "GROUP BY AG.AGENT_NUMBER, AGENT_NAME, AGENT_STATUS, AE ORDER BY " + Orderby + " ";



                SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);

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
                regioncode.Text = sRegion;

            }

            //grdHandling.DataSource = datatab;
            //grdHandling.DataBind();

            // adPolicy.Fill(dsPolicy);

            con.Close();
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
            //if (DataRowCurrView["REGION_NAME"] != null && DataRowCurrView["REGION_NAME"].ToString().Trim() != "")
            //    strRowValue[0] +=DataRowCurrView["REGION_NAME"] + "~";
            //if (DataRowCurrView["REGION_CODE"] != null && DataRowCurrView["REGION_CODE"].ToString().Trim() != "")
            //    strRowValue[1] += DataRowCurrView["REGION_CODE"] + "~";
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
            grdHandling.DataSource = datatab;
            grdHandling.DataBind();

        }

        protected void grdHandling_RowDataBound(object sender, GridViewRowEventArgs e)
        {
             string sCompany = Request.QueryString["COMPANY_CODE"];
            string sRegion = (string)Request.QueryString["Region_Code"];
            string sFromDate1 = (string)Request.QueryString["Fromdate"];
            string sToDate1 = (string)Request.QueryString["Today"];
            string sRider = (string)Request.QueryString["bType"];

            if (e.Row.RowType == DataControlRowType.DataRow)
            {

                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();

                string sAgent = e.Row.Cells[0].Text;
                //   string Policy_num = e.Row.Cells[2].Text;
                if ((e.Row.RowType == DataControlRowType.DataRow) || (e.Row.RowType == DataControlRowType.Header))
                {
                    e.Row.Cells[e.Row.Cells.Count - 1].Visible = false;
                    grdHandling.HeaderRow.Cells[e.Row.Cells.Count - 1].Visible = false;

                }
                e.Row.Cells[1].ToolTip = "click to view details";

                string text = e.Row.Cells[0].Text;
                HyperLink link = new HyperLink();
                link.NavigateUrl = "Claimsforagent.aspx?sAgent=" + sAgent + "&Fromdate=" + sFromDate1 + "&Today=" + sToDate1 + "&bType=" + sRider + "&COMPANY_CODE=" + sCompany + "&Region_Code=" + sRegion + "";
                link.Text = text;
                link.Target = "_blank";
                e.Row.Cells[0].Controls.Add(link);

                //PolicyView.aspx?POLICY_NUMBER=W861860006&COMPANY_CODE=16&AGENT_NUMBER=ID01001
                //e.Row.Cells[2].Text = Convert.ToString("<a href=\"PolicyView.aspx?POLICY_NUMBER="+Policy_num+"&COMPANY_CODE="+sCompany+"&AGENT_NUMBER="+Agent_num+"Target="+"_blank"+" \">"+Policy_num+"</a>");
            }

        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            string sCompany = Request.QueryString["COMPANY_CODE"];
            string sRegion = (string)Request.QueryString["Region_Code"];
            string sFromDate1 = (string)Request.QueryString["Fromdate"];
            string sToDate1 = (string)Request.QueryString["Today"];
            string sRider = (string)Request.QueryString["bType"];
            Response.Redirect("ClaimsByRegion.aspx?Region_Code=" + sRegion + "&Fromdate=" + sFromDate1 + "&Today=" + sToDate1 + "&bType=" + sRider + "&COMPANY_CODE=" + sCompany + "");

        }
        protected void Button2_Click(object sender, EventArgs e)
        {
            //Export the grid data to excel without hitting the Database
            ExportToExcel();
            //Export("ClaimsByRegion.xls", this.grRequirementsByRegion);
        }




        protected void ExportToExcel()
        {
            //InvokeSP();
            //            dataPolicy = Session["dataPolicy"] as DataTable;

            //dataPolicy = GridView1.DataSource as DataTable;
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                dataPolicy.DefaultView.Sort = "REGION_NAME";
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