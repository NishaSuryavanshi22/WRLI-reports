using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CSCUtils;
using System.IO;
using System.Globalization;

namespace WRLI_Reports
{
    public partial class Ratedforagent : System.Web.UI.Page
    {
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);


        DataSet dsPolicy = new DataSet();
        public static DataTable dataPolicy = new DataTable();
        string sCompany = "15";
        string sRegionCode = "INS";
        string resulttype="ALL";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = "PR.POLICY_NUMBER ASC";
        string OrderDir = "ASC";
        bool bNet = false;
        bool bRegion = false;
        int nRowct = 0;
        string sType = "type";
        Int32[] arr_NB = new Int32[] { };
        string[] strRowValue = new string[3];
        DataTable datatab = new DataTable(); // Create a new Data table
        DataTable datatabTotal = new DataTable();

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

            string sRateDate = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            string sPaidColumn = "PAYMENT_DATE";
            string sCompany = Request.QueryString["COMPANY_CODE"];
            string sAgentid = (string)Request.QueryString["Agentid"];
            string sFromDate = (string)Request.QueryString["FromDate"];
            string sToDate = (string)Request.QueryString["Today"];
            string sState = (string)Request.QueryString["sState"];
            string StausQuery = (string)Request.QueryString["np=1"];


            SqlCommand commPolicy = new SqlCommand();

            commPolicy.Connection = con;
            //commPolicy.CommandType = CommandType.StoredProcedure;

            commPolicy.CommandText = "SELECT Distinct(PR.POLICY_NUMBER), PR.PI_STATE,PR.APP_RECEIVED_DATE,PR.PAYMENT_FLAG,PR.PI_LAST, RTRIM(PR.PI_LAST)+', '+RTRIM(PR.PI_FIRST) AS PI_NAME,PR.PI_BUSINESS,PR.PLAN_CODE,PR.RATE_CLASS,(case when ((PB1.SL_PERCENT > 0) AND (PB1.SL_PCT_CEASE_DATE>" + sRateDate + ")) THEN PB1.SL_PERCENT ELSE 0 END) AS PERCENT_RATING, (case when ((PB1.SL_FLAT_AMOUNT > 0) AND (PB1.SL_FLAT_CEASE_DATE>" + sRateDate + ")) THEN PB1.SL_FLAT_AMOUNT ELSE 0 END) AS FLAT_AMOUNT_RATING, PR.PAYMENT_DATE, PR.CONTRACT_CODE, ISNULL(PR.CONTRACT_REASON,'') AS CONTRACT_REASON, CASE WHEN (PR.CONTRACT_CODE <> 'T') THEN '' ELSE PR.CONTRACT_DESC END AS CANCEL_REASON,(PR.ANNUAL_PREMIUM) * (CASE WHEN (ST.PROD_PCNT IS NULL) THEN 1 ELSE (ST.PROD_PCNT/100) END) as ANNUALIZED_PREM, ISNULL(ST.PROD_PCNT,100) AS PROD_PCNT, (CASE WHEN (ST.SERVICE_AGENT_IND <> 'X') THEN '*' ELSE '' END) as TRANSFER_IND,PR.RECORD_TYPE FROM POLICIES AS PR WITH (NOLOCK) LEFT OUTER JOIN POLICY_COVERAGE AS PB1 WITH (NOLOCK) ON PB1.COMPANY_CODE = PR.COMPANY_CODE AND PB1.POLICY_NUMBER = PR.POLICY_NUMBER AND PB1.BENEFIT_TYPE = 'SL' LEFT OUTER JOIN COVERAGE2 C2 WITH (NOLOCK) ON C2.COVERAGE_ID = PR.PRODUCT_CODE LEFT OUTER JOIN POLICY_SPLIT ST WITH (NOLOCK) ON (PR.COMPANY_CODE = ST.COMPANY_CODE) and (PR.POLICY_NUMBER = ST.POLICY_NUMBER) AND (PR.AGENT_NUMBER = ST.AGENT_NUMBER) and (PR.ISSUE_DATE = ST.ISSUE_DATE) WHERE PR.AGENT_NUMBER = '" + sAgentid + "'  AND PR.AGENT_NUMBER IN (SELECT AH.AGENT_NUMBER FROM AGENT_HIERLIST AH WHERE AH.HIERARCHY_AGENT = '" + AgentID+  "') AND ((PR.APP_RECEIVED_DATE BETWEEN '" + sFromDate + "' AND '" + sToDate+"') OR (PR.PAYMENT_DATE BETWEEN '" + sFromDate+"' AND  '"+ sToDate + "')) AND  ((C2.MED_TYPE = '" + resulttype+"') OR ('" + resulttype+ "' = 'ALL'))   AND PR.COMPANY_CODE = '" + sCompany+ "'  AND (('" + sState + "'='ALL')OR(PR.PI_STATE = '" + sState + "')) ORDER BY " + Orderby;
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
                    grdratedforagent.Visible = false;

                }
            }
            else
            {
                int nTotCt = datatab.Rows.Count;
                grdratedforagent.Visible = true;
                lblcount.Visible = true;
                //
                Int32 sNullCheck = 0;
                DataView DataView1 = new DataView(datatab);
                DataRowView DataFirstRowView = null;
                DataFirstRowView = DataView1[0];

                // sNullCheck = Convert.ToInt32(DataFirstRowView["SUB_COUNT"]);
                //if (string.IsNullOrEmpty(sNullCheck) || sNullCheck == "0")
                //if (sNullCheck == 0)
                //{
                //   lblcount.Text = "No Records Found for the selected criteria !!!";
                //   dvgrid.Visible = false;

                //    return;
                //  }
                //AddEditRows();
                lblcount.Text = "Total Record Count: " + nTotCt;
                Agent_Number.Text = sAgentid;

;

            }
            grdratedforagent.DataSource = datatab;
            grdratedforagent.DataBind();

            // adPolicy.Fill(dsPolicy);

            con.Close();

        }
        protected void grdratedforagent_RowDataBound(object sender, GridViewRowEventArgs e)

        {
            string sCompany = Request.QueryString["COMPANY_CODE"];
            string sRegion = (string)Request.QueryString["Region_Code"];
            string sFromDate1 = (string)Request.QueryString["Fromdate"];
            string sToDate1 = (string)Request.QueryString["Today"];
            string sState = Request.QueryString["State"];
            string StausQuery = "&STATE=" + sState;
            string sAgentid = (string)Request.QueryString["Agentid"];

            //string sState = "ALL";
            //string[] fromDate;
            //string[] toDate;
            //if (txtFrom.Text.Contains('/'))
            //{
            //    fromDate = txtFrom.Text.Split('/');
            //    FromDate = fromDate[2] + fromDate[0] + fromDate[1];
            //    //
            //}
            //else if (txtFrom.Text.Contains('-'))
            //{
            //    fromDate = txtFrom.Text.Split('-');
            //    FromDate = fromDate[2] + fromDate[0] + fromDate[1];
            //}
            //if (txtTo.Text.Contains('/'))
            //{
            //    toDate = txtTo.Text.Split('/');
            //    ToDate = toDate[2] + toDate[0] + toDate[1];
            //}

            //else if (txtTo.Text.Contains('-'))
            //{
            //    toDate = txtTo.Text.Split('-');
            //    ToDate = toDate[2] + toDate[0] + toDate[1];
            //}


            if (e.Row.RowType == DataControlRowType.DataRow)
            {

                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();

                string Policy_number = e.Row.Cells[0].Text;
                //   string Policy_num = e.Row.Cells[2].Text;
                if ((e.Row.RowType == DataControlRowType.DataRow) || (e.Row.RowType == DataControlRowType.Header))
                {
                    e.Row.Cells[e.Row.Cells.Count - 1].Visible = false;
                    grdratedforagent.HeaderRow.Cells[e.Row.Cells.Count - 1].Visible = false;

                }
                e.Row.Cells[0].ToolTip = "click to view details";

                string text = e.Row.Cells[0].Text;
                HyperLink link = new HyperLink();
                link.NavigateUrl = "PolicyView.aspx?POLICY_NUMBER=" + Policy_number + "&COMPANY_CODE=" + sCompany + "&AGENT_NUMBER=" + sAgentid + "";
                link.Text = text;
                link.Target = "_blank";
                e.Row.Cells[0].Controls.Add(link);

                //PolicyView.aspx?POLICY_NUMBER=W861860006&COMPANY_CODE=16&AGENT_NUMBER=ID01001
                //e.Row.Cells[2].Text = Convert.ToString("<a href=\"PolicyView.aspx?POLICY_NUMBER="+Policy_num+"&COMPANY_CODE="+sCompany+"&AGENT_NUMBER="+Agent_num+"Target="+"_blank"+" \">"+Policy_num+"</a>");
            }

        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            string sPaidColumn = "PAYMENT_DATE";
            string sCompany = Request.QueryString["COMPANY_CODE"];
            string sRegion = (string)Request.QueryString["Region_Code"];
            string sFromDate = (string)Request.QueryString["Fromdate"];
            string sToDate = (string)Request.QueryString["Today"];
            string sState = (string)Request.QueryString["sState"];
            Response.Redirect("RatedforRegion.aspx?Region_Code=" + sRegion + "&Fromdate=" + sFromDate + "&Today=" + sToDate + "&sState=" + sState + "&COMPANY_CODE=" + sCompany + "");

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