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
using System.Configuration;

namespace WRLI_Reports
{
    public partial class HandlingTimeReport : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //        SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 1000");

        // SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        string agent = "WRE";
        private string frmDate;
        private string tDate;
        DataSet dsPolicy = new DataSet();
        public static DataTable dataPolicy = new DataTable();
   

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                tblgrid.Visible = false;
                txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                Session["CompanyCode"] = "15";
                //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Initial Catalog=WRE_AGENT");
                SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 1000");
                con.Open();

                SqlCommand commCompany = new SqlCommand("select * from company_details ");
                commCompany.Connection = con;
                DataSet dscomp = new DataSet();
                SqlDataAdapter adcomp = new SqlDataAdapter(commCompany.CommandText, con);
                adcomp.Fill(dscomp);
                List<string> lstcomp = new List<string>();
                for (int i = 0; i < dscomp.Tables[0].Rows.Count; i++)
                {
                    lstcomp.Add(dscomp.Tables[0].Rows[i].ItemArray[0].ToString() + "-" + dscomp.Tables[0].Rows[i].ItemArray[1].ToString());

                }
                ddlHandcompany.DataSource = lstcomp;
                ddlHandcompany.DataBind();

                SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + Session["CompanyCode"] + "' AND HIERARCHY_AGENT = '" + agent + "' ORDER BY SORTNAME ASC");
                comm.Connection = con;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(comm.CommandText, con);
                ad.Fill(ds);

                List<string> lstagent = new List<string>();
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    lstagent.Add(ds.Tables[0].Rows[i].ItemArray[0].ToString() + "-" + ds.Tables[0].Rows[i].ItemArray[1].ToString());

                }

                ddlHandagent.DataSource = lstagent;
                ddlHandagent.DataBind();

                SqlCommand commStates = new SqlCommand("SELECT * FROM ALLSTATES ORDER BY STATE_ABBR ASC");
                commStates.Connection = con;
                DataSet dsStates = new DataSet();
                SqlDataAdapter adStates = new SqlDataAdapter(commStates.CommandText, con);
                adStates.Fill(dsStates);
                ddlHandstate.DataSource = dsStates;
                ddlHandstate.DataTextField = "STATE_NAME";
                ddlHandstate.DataValueField = "STATE_ABBR";
                ddlHandstate.DataBind();

                SqlCommand commReason = new SqlCommand("SELECT DISTINCT CONTRACT_REASON, UPPER(CONTRACT_DESC) AS CONTRACT_DESC FROM POLICIES2 WHERE RTRIM(ISNULL(CONTRACT_REASON,''))<>'' ORDER BY UPPER(CONTRACT_DESC)");
                commReason.Connection = con;
                DataSet dsReason = new DataSet();
                SqlDataAdapter adReason = new SqlDataAdapter(commReason.CommandText, con);
                adReason.Fill(dsReason);
                ddlHandpolicydesc.DataSource = dsReason;
                ddlHandpolicydesc.DataTextField = "CONTRACT_DESC";
                ddlHandpolicydesc.DataValueField = "CONTRACT_REASON";
                ddlHandpolicydesc.DataBind();

                //SqlCommand commMarket = new SqlCommand("SELECT ISNULL(MARKETING_COMPANY,'UNKNOWN') AS MARKETING_COMPANY FROM REGION_NAMES WHERE MARKETING_COMPANY LIKE '1N%' ORDER BY MARKETING_COMPANY ASC");
                SqlCommand commMarket = new SqlCommand("SELECT DISTINCT ISNULL(MARKETING_COMPANY,'UNKNOWN') AS MARKETING_COMPANY, CASE WHEN AGENCY_NAME IS NULL THEN AGENT_NAME ELSE AGENCY_NAME END AS " + " [REGION_NAME] FROM REGION_NAMES WHERE MARKETING_COMPANY LIKE '1N%' OR MARKETING_COMPANY LIKE 'C1%' OR MARKETING_COMPANY LIKE 'INS%' ORDER BY MARKETING_COMPANY ASC");
                commMarket.Connection = con;
                DataSet dsMarket = new DataSet();
                SqlDataAdapter adMarket = new SqlDataAdapter(commMarket.CommandText, con);
                adMarket.Fill(dsMarket);
                ddlHandregion.DataSource = dsMarket;
                ddlHandregion.DataTextField = "MARKETING_COMPANY";
                ddlHandregion.DataBind();

                con.Close();
            }
        }




        protected void Button1_Click(object sender, EventArgs e)
        {
            tblgrid.Visible = true;
            string selectedComp = "ALL";
            string selectedAgent = "ALL";
            string sInforce;
            string sDateRange;

            string[] fromDate;
            string[] toDate;
           
            if (txtFrom.Text.Contains('/'))
            {
                fromDate = txtFrom.Text.Split('/');
               frmDate = fromDate[2] + fromDate[0] + fromDate[1];
                //
            }
            else if (txtFrom.Text.Contains('-'))
            {
                fromDate = txtFrom.Text.Split('-');
                 frmDate = fromDate[2] + fromDate[0] + fromDate[1];
            }
            if (txtTo.Text.Contains('/'))
            {
                toDate = txtTo.Text.Split('/');
                 tDate = toDate[2] + toDate[0] + toDate[1];
            }

            else if (txtTo.Text.Contains('-'))
            {
                toDate = txtTo.Text.Split('-');
               tDate = toDate[2] + toDate[0] + toDate[1];
            }

            //string[] fromDate = txtFrom.Text.Split('/');
            //string frmDate = fromDate[2] + fromDate[0] + fromDate[1];

            //string[] toDate = txtTo.Text.Split('/');
            //string tDate = toDate[2] + toDate[0] + toDate[1];
            SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 1000");


            con.Open();

            int index = ddlHandcompany.SelectedItem.Value.LastIndexOf("-");
            if (index > 0)
            {
                selectedComp = ddlHandcompany.SelectedItem.Value.Substring(0, index);
            }

            int indexagent = ddlHandagent.SelectedItem.Value.LastIndexOf("-");
            if (indexagent > 0)
            {
                selectedAgent = ddlHandagent.SelectedItem.Value.Substring(0, indexagent);
            }
            //int indexagentnum = ddlHandagent.SelectedItem.Value.LastIndexOf("- -");
            //if (indexagentnum > 0)
            //{
            //    selectedComp = ddlHandcompany.SelectedItem.Value.Substring(0, index);
            //}

            sInforce = "";
            if(ddlHanddatatype.SelectedValue== "SUBMITTED")
            {
                sDateRange = "po.APP_RECEIVED_DATE";
            }
            else
            {
                sInforce = " AND (PO.RECORD_TYPE = 'I') ";
                sDateRange = "po.ISSUE_DATE";
            }


            /* SqlCommand commPolicy = new SqlCommand("select dbo.GET_POLICY_INSURED(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS INSURED_NAME, RIGHT(RTRIM(ISNULL(PI_SOC_SEC_NUMBER,'0000')),4) AS INSURED_SSN,dbo.GET_POLICY_OWNER(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS OWNER_NAME,dbo.GET_POLICY_PAYOR(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS PAYOR_NAME,PO.COMPANY_CODE,PO.POLICY_NUMBER, dbo.[GET_BASE_PRODUCT_DESCEX](ISNULL(PO.PRODUCT_CODE,''),ISNULL(PP.PRODUCT_CODE,'')) AS PLAN_TYPE,po.RATE_CLASS,dbo.GET_POLICY_ISSUE_STATE(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS ISSUE_STATE,dbo.GET_POLICY_PI_CITY(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS INSURED_CITY, ISNULL(PI_STATE,dbo.GET_POLICY_PI_STATE(PO.COMPANY_CODE,PO.POLICY_NUMBER)) AS INSURED_STATE,dbo.LPDATE_TO_STRDATE(ISNULL(dbo.GET_POLICY_PI_DOB(PO.COMPANY_CODE,PO.POLICY_NUMBER),'')) AS INSURED_DOB,dbo.GET_POLICY_ISSUE_AGE(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS ISSUE_AGE,dbo.FORMAT_PHONE(ISNULL(dbo.GET_POLICY_PI_PHONE(PO.COMPANY_CODE,PO.POLICY_NUMBER),'0000000000')) AS INSURED_PHONE,ISNULL(po.FACE_AMOUNT,0) AS FACE_AMOUNT,ISNULL(pp.MODE_PREMIUM,'0') AS MODE_PREMIUM,UPPER(dbo.GETTRANSLATION('POLICY INFO:BILLING_MODE',RTRIM(CASE WHEN (ISNULL(BILLING_MODE,'')<10) THEN '0'+RTRIM(ISNULL(BILLING_MODE,'')) ELSE RTRIM(ISNULL(BILLING_MODE,'')) END))) AS BILLING_MODE,dbo.GET_POLICY_BILLING_FORM(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS BILLING_FORM,PO.CONTRACT_CODE AS STATUS,CASE WHEN (PO.CONTRACT_CODE = 'A') THEN 'ACTIVE' ELSE PO.CONTRACT_DESC END AS STATUS_DESC,dbo.GET_AGENT_DISPLAY_NAME(PO.COMPANY_CODE,PO.AGENT_NUMBER,'L') AS SERVICE_AGENT_NAME,PO.AGENT_NUMBER AS SERVICE_AGENT,dbo.GET_AGENT_CO_PHONE(PO.COMPANY_CODE,PO.AGENT_NUMBER) AS SERVICE_AGENT_PHONE,dbo.GET_AGENT_DISPLAY_NAME(PO.COMPANY_CODE,ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER),'L') AS WRITING_AGENT_NAME,ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER) AS WRITING_AGENT,dbo.GET_AGENT_CO_PHONE(PO.COMPANY_CODE,ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER)) AS WRITING_AGENT_PHONE,PO.DURATION AS DURATION,dbo.LPDATE_TO_STRDATE(ISNULL(PO.ISSUE_DATE,'')) AS ISSUE_DATE,dbo.LPDATE_TO_STRDATE(ISNULL(PO.PAID_TO_DATE,'')) AS PAID_TO_DATE, 0 AS CASH_VALUE,PO.RATE_CLASS,dbo.GET_AGENT_STATUS(PO.COMPANY_CODE,PO.AGENT_NUMBER) AS SA_STATUS,dbo.GET_AGENT_STATUS(PO.COMPANY_CODE,ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER)) AS WA_STATUS,CASE WHEN (ISNULL(PO.REGION_CODE,'0')='0') THEN dbo.GETAGENTREGION(PO.AGENT_NUMBER,PO.COMPANY_CODE) ELSE PO.REGION_CODE END AS SA_REGION_CODE,dbo.GETAGENTREGION(ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER),PO.COMPANY_CODE) AS WA_REGION_CODE, dbo.LPDATE_TO_STRDATE(application_date) as application_date, dbo.LPDATE_TO_STRDATE(po.app_received_date) as app_received_date, case when (po.contract_code='T') then dbo.LPDATE_TO_STRDATE(po.last_change_date) else dbo.LPDATE_TO_STRDATE('') end as termination_date,ISNULL(po.annual_premium,0) as ANNUAL_PREMIUM,dbo.GET_POLICY_PI_GENDER(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS Insured_gender,ISNULL(PO.ANNUAL_PREMIUM,0) as ANNUAL_PREMIUM from POLICIES2 PO left outer join pending_policy pp on po.policy_number = pp.policy_number and po.company_code=pp.company_code WHERE (PO.AGENT_NUMBER IN (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = '" + agent + "' and company_code = '" + selectedComp + "' OR '15'='ALL')) AND (po.APP_RECEIVED_DATE BETWEEN '" + frmDate + "' and '" + tDate + "' ) AND (PO.COMPANY_CODE = '" + Session["CompanyCode"] + "' OR '15'='ALL') AND (PO.AGENT_NUMBER = '" + selectedAgent + "' OR 'ALL'='ALL') AND (PI_STATE = '" + ddlstate.SelectedItem.Value + "' OR 'ALL' = 'ALL')AND (po.CONTRACT_CODE = '" + ddlpolicystatus.SelectedItem.Value + "' OR 'ALL' = 'ALL') AND (CONTRACT_REASON = '" + ddlpolicydesc.SelectedItem.Value + "' OR 'ALL' = 'ALL') AND (PO.REGION_CODE = '" + ddlregion.SelectedItem.Value + "' OR 'ALL' = 'ALL') ORDER BY POLICY_NUMBER DESC");  */

            /* SqlCommand commPolicy = new SqlCommand("select dbo.GET_POLICY_INSURED(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS INSURED_NAME, PO.COMPANY_CODE,PO.POLICY_NUMBER,PO.REGION_CODE, dbo.[GET_BASE_PRODUCT_DESCEX](ISNULL(PO.PRODUCT_CODE,''),ISNULL(PP.PRODUCT_CODE,'')) AS PLAN_TYPE,dbo.GET_POLICY_ISSUE_STATE(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS ISSUE_STATE,dbo.GET_POLICY_PI_CITY(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS INSURED_CITY,PO.CONTRACT_CODE AS STATUS,CASE WHEN (PO.CONTRACT_CODE = 'A') THEN 'ACTIVE' ELSE PO.CONTRACT_DESC END AS STATUS_DESC,dbo.LPDATE_TO_STRDATE(ISNULL(PO.ISSUE_DATE,'')) AS ISSUE_DATE,dbo.LPDATE_TO_STRDATE(ISNULL(PO.PAID_TO_DATE,'')) AS PAID_TO_DATE, dbo.LPDATE_TO_STRDATE(application_date) as application_date, dbo.LPDATE_TO_STRDATE(po.app_received_date) as app_received_date from POLICIES2 PO left outer join pending_policy pp on po.policy_number = pp.policy_number and po.company_code=pp.company_code WHERE (PO.AGENT_NUMBER IN (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = '" + agent + "' and company_code = '" + selectedComp + "' OR '15'='ALL')) AND (po.APP_RECEIVED_DATE BETWEEN '" + frmDate + "' and '" + tDate + "' ) AND (PO.COMPANY_CODE = '" + Session["CompanyCode"] + "' OR '15'='ALL') AND (PO.AGENT_NUMBER = '" + selectedAgent + "' OR 'ALL'='ALL') AND (PI_STATE = '" + ddlstate.SelectedItem.Value + "' OR 'ALL' = 'ALL')AND (po.CONTRACT_CODE = '" + ddlpolicystatus.SelectedItem.Value + "' OR 'ALL' = 'ALL') AND (CONTRACT_REASON = '" + ddlpolicydesc.SelectedItem.Value + "' OR 'ALL' = 'ALL') AND (PO.REGION_CODE = '" + ddlregion.SelectedItem.Value + "' OR 'ALL' = 'ALL') ORDER BY POLICY_NUMBER DESC"); */

            //For Testing, Later use below query
           SqlCommand commPolicy = new SqlCommand("select dbo.GET_POLICY_INSURED(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS Insured_Name, PO.Company_Code,PO.Policy_Number,PO.Region_Code, dbo.[GET_BASE_PRODUCT_DESCEX](ISNULL(PO.PRODUCT_CODE,''),ISNULL(PP.PRODUCT_CODE,'')) AS Plan_Type,dbo.GET_POLICY_ISSUE_STATE(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS Issue_State,dbo.GET_POLICY_PI_CITY(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS Insured_City,PO.CONTRACT_CODE AS Status,CASE WHEN (PO.CONTRACT_CODE = 'A') THEN 'ACTIVE' ELSE PO.CONTRACT_DESC END AS Status_Desc,dbo.LPDATE_TO_STRDATE(ISNULL(PO.ISSUE_DATE,'')) AS Issue_Date,dbo.LPDATE_TO_STRDATE(ISNULL(PO.PAID_TO_DATE,'')) AS Paid_to_date, dbo.LPDATE_TO_STRDATE(application_date) as Application_date, dbo.LPDATE_TO_STRDATE(po.app_received_date) as App_received_date from POLICIES2 PO left outer join pending_policy pp on po.policy_number = pp.policy_number and po.company_code=pp.company_code WHERE (PO.AGENT_NUMBER IN (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = '" + agent + "' and company_code = '" + selectedComp + "' OR '15'='ALL')) AND ("+sDateRange+" BETWEEN '" + frmDate + "' and '" + tDate + "' ) "+sInforce+" AND (PO.COMPANY_CODE = '" + Session["CompanyCode"] + "' OR '15'='ALL') AND (PO.AGENT_NUMBER = '" + selectedAgent + "' OR 'ALL'='ALL') AND (PI_STATE = '" + ddlHandstate.SelectedItem.Value + "' OR 'ALL' = 'ALL')AND (po.CONTRACT_CODE = '" + ddlHandpolicystatus.SelectedItem.Value + "' OR 'ALL' = 'ALL') AND (CONTRACT_REASON = '" + ddlHandpolicydesc.SelectedItem.Value + "' OR 'ALL' = 'ALL') AND (PO.REGION_CODE = '" + ddlHandregion.SelectedItem.Value + "' OR 'ALL' = 'ALL') ORDER BY POLICY_NUMBER DESC");
           
         //   SqlCommand commPolicy = new SqlCommand("select dbo.GET_POLICY_INSURED(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS Insured_Name, PO.Company_Code,PO.Policy_Number,PO.Region_Code, dbo.[GET_BASE_PRODUCT_DESCEX](ISNULL(PO.PRODUCT_CODE,''),ISNULL(PP.PRODUCT_CODE,'')) AS Plan_Type,dbo.GET_POLICY_ISSUE_STATE(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS Issue_State,dbo.GET_POLICY_PI_CITY(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS Insured_City,PO.CONTRACT_CODE AS Status,CASE WHEN (PO.CONTRACT_CODE = 'A') THEN 'ACTIVE' ELSE PO.CONTRACT_DESC END AS Status_Desc,dbo.LPDATE_TO_STRDATE(ISNULL(PO.ISSUE_DATE,'')) AS Issue_Date,dbo.LPDATE_TO_STRDATE(ISNULL(PO.PAID_TO_DATE,'')) AS Paid_to_date, dbo.LPDATE_TO_STRDATE(application_date) as Application_date, dbo.LPDATE_TO_STRDATE(po.app_received_date) as App_received_date from POLICIES2 PO left outer join pending_policy pp on po.policy_number = pp.policy_number and po.company_code=pp.company_code WHERE (PO.AGENT_NUMBER IN (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = '" + agent + "' and company_code = '" + selectedComp + "' OR '15'='ALL')) AND (" + sDateRange + " BETWEEN '" + frmDate + "' and '" + tDate + "' ) " + sInforce + " AND (PO.COMPANY_CODE = '" + Session["CompanyCode"] + "' OR '15'='ALL') AND (PO.AGENT_NUMBER = '" + selectedAgent + "' OR '"+selectedAgent+"'='ALL') AND (PI_STATE = '" + ddlHandstate.SelectedItem.Value + "' OR '" + ddlHandstate.SelectedItem.Value + "' = 'ALL')AND (po.CONTRACT_CODE = '" + ddlHandpolicystatus.SelectedItem.Value + "' OR '" + ddlHandpolicystatus.SelectedItem.Value + "' = 'ALL') AND (CONTRACT_REASON = '" + ddlHandpolicydesc.SelectedItem.Value + "' OR '" + ddlHandpolicydesc.SelectedItem.Value + "' = 'ALL') AND (PO.REGION_CODE = '" + ddlHandregion.SelectedItem.Value + "' OR '" + ddlHandregion.SelectedItem.Value + "' = 'ALL') ORDER BY POLICY_NUMBER DESC");
            commPolicy.Connection = con;

            SqlDataAdapter adPolicy = new SqlDataAdapter(commPolicy.CommandText, con);
            adPolicy.Fill(dsPolicy);
            con.Close();

            if (dsPolicy.Tables[0].Rows.Count == 0)
            {
                DataTable dt = new DataTable();
                DataColumn dc = new DataColumn();

                if (dt.Columns.Count == 0)
                {
                    dt.Columns.Add("INSURED NAME", typeof(string));

                    dt.Columns.Add("Region Code", typeof(string));
                    dt.Columns.Add("COMPANY CODE", typeof(string));
                    dt.Columns.Add("POLICY NUMBER", typeof(string));
                    dt.Columns.Add("PLAN", typeof(string));
                    dt.Columns.Add("AGENT NAME", typeof(string));
                    dt.Columns.Add("AGENT NUMBER", typeof(string));

                    dt.Columns.Add("ISSUE STATE", typeof(string));
                    dt.Columns.Add("INSURED CITY", typeof(string));

                    // dt.Columns.Add("FACE AMT/MO.INC", typeof(string));
                    dt.Columns.Add("STATUS", typeof(string));
                    dt.Columns.Add("STATUS DESC", typeof(string));

                    dt.Columns.Add("APPLICATION SIGNED DATE", typeof(string));
                    dt.Columns.Add("APPLICATION RECEIVED DATE", typeof(string));
                    //dt.Columns.Add("DURATION", typeof(string));
                    dt.Columns.Add("ISSUE DATE", typeof(string));
                    dt.Columns.Add("PAID TO DATE", typeof(string));


                }
                dvgrid.Style.Add("height", "120px");
                DataRow NewRow = dt.NewRow();
                dt.Rows.Add(NewRow);
                grdHandling.DataSource = dt;
                grdHandling.DataBind();
                lblcount.Text = "No Records Found for the selected criteria !!";

            }
            else
            {
                dvgrid.Style.Add("height", "600px");
                grdHandling.DataSource = dsPolicy;
                grdHandling.DataBind();
                lblcount.Text = dsPolicy.Tables[0].Rows.Count.ToString();
                dataPolicy = dsPolicy.Tables[0];
            }
        }

        protected void ddlagent_PreRender(object sender, EventArgs e)
        {
            if (ddlHandagent.Items.FindByValue(string.Empty) == null)
            {
                ddlHandagent.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
        }

        protected void ddlregion_PreRender(object sender, EventArgs e)
        {
            if (ddlHandregion.Items.FindByValue(string.Empty) == null)
            {
                ddlHandregion.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
        }


        protected void ddlstate_PreRender(object sender, EventArgs e)
        {
            if (ddlHandstate.Items.FindByValue(string.Empty) == null)
            {
                ddlHandstate.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
        }


        protected void ddlpolicydesc_PreRender(object sender, EventArgs e)
        {
            if (ddlHandpolicydesc.Items.FindByValue(string.Empty) == null)
            {
                ddlHandpolicydesc.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            //Export("PolicyReport.xls", this.grdHandling);
            // New changes for Report 28/02/2017
            ExportToExcel();

        }

        protected void ExportToExcel()
        {
            //InvokeSP();
            //            dataPolicy = Session["dataPolicy"] as DataTable;
            dataPolicy.DefaultView.Sort = "POLICY_NUMBER";
            dataPolicy = dataPolicy.DefaultView.ToTable();
            //dataPolicy = GridView1.DataSource as DataTable;
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"Policy_Report_" + DateTime.Now.ToString("yyyyMMdd");
                Response.AddHeader("Content-Disposition", "inline;filename=" + filename.Replace("/", "").Replace(":", "") + ".xlsx");
                //Call  Export function
                Response.BinaryWrite(Utils.ExportToCSVFileOpenXML(dataPolicy));

                Response.Flush();
                Response.End();
            }

        }

        public static void Export(string fileName, GridView gv)
        {
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.AddHeader(
                "content-disposition", string.Format("attachment; filename={0}", fileName));
            HttpContext.Current.Response.ContentType = "application/ms-excel";

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
