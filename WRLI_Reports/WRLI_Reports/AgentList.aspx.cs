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
    public partial class AgentList : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        // SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
      //  SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 1000");
        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        DataSet dsPolicy = new DataSet();

        DataTable datatab = new DataTable(); // Create a new Data table
        public static DataTable dataPolicy = new DataTable();

        string sCompany = "05";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = "REGION_CODE";
        string OrderDir = "ASC";
        string sState = "ALL";
        string sStatus = "ALL";
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                tblgrid.Visible = false;
                //txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                //txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                        AgentID = Session["LoginID"].ToString();

                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();

                //Session["CompanyCode"] = "05";
                //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Initial Catalog=WRE_AGENT");
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

                /*SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + AgentID + "' ORDER BY SORTNAME ASC");
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
                ddlHandagent.DataBind(); */

                con.Close();
            }
        }




        protected void Button1_Click(object sender, EventArgs e)
        {
            tblgrid.Visible = true;
            string selectedComp = "ALL";
            string selectedAgent = "ALL";
            //string[] fromDate = txtFrom.Text.Split('/');
            //string frmDate = fromDate[2] + fromDate[0] + fromDate[1];

            //string[] toDate = txtTo.Text.Split('/');
            //string tDate = toDate[2] + toDate[0] + toDate[1];

            //if(ddlHandcompany.SelectedValue != null )
               // sCompany = ddlHandcompany.SelectedValue;
            con.Open();

            int index = ddlHandcompany.SelectedItem.Value.LastIndexOf("-");
            if (index > 0)
            {
                sCompany = ddlHandcompany.SelectedItem.Value.Substring(0, index);
            } 

            /*int indexstate = ddlHandstate.SelectedItem.Value.LastIndexOf("-");
            if (indexstate > 0)
            {
                sState = ddlHandstate.SelectedItem.Value.Substring(0, indexstate);
            } */

            int indexStatus = ddlHandpolicystatus.SelectedItem.Value.LastIndexOf("-");
            if (indexStatus > 0)
            {
                //sStatus = ddlHandpolicystatus.SelectedItem.Value.Substring(0, indexStatus);
            }
            if(ddlHandpolicystatus.SelectedItem.Value != null)
            {
                sStatus = ddlHandpolicystatus.SelectedItem.Value;
                if (sStatus == "ACTIVE")
                {
                    sStatus = "A";
                }
                if (sStatus == "TERMINATED")
                { 
                    sStatus = "T";
                }
            }

            //New Query
            SqlCommand commPolicy = new SqlCommand("SELECT COMPANY_CODE, A.STATUS_CODE, A.AGENT_NUMBER,CASE WHEN (NAME_FORMAT_CODE = 'C') THEN RTRIM(NAME_BUSINESS) ELSE RTRIM(INDIVIDUAL_LAST)+', '+RTRIM(INDIVIDUAL_FIRST) END AS AGENT_NAME,  NAME_FORMAT_CODE, REGION_CODE,CASE WHEN (ISNULL(AGENCY_NAME,'') ='') THEN AGENT_NAME ELSE AGENCY_NAME END AS REGION_NAME, A.DEAL_CODE, ISNULL(A.eMail,'') AS EMAIL, ISNULL(START_DATE,'') AS CONTRACT_DATE,ISNULL(STOP_DATE,'') AS STOP_DATE,LICENSED_STATE, CASE WHEN (RTRIM(ISNULL(TELE_NUM_OFFICE,''))<>'') THEN dbo.FORMAT_PHONE(RTRIM(TELE_NUM_OFFICE)) ELSE '' END AS AGENT_PHONE, CASE WHEN (RTRIM(ISNULL(ADDR_LINE_1,''))<>'') THEN  RTRIM(ADDR_LINE_1) + ', ' ELSE '' END + CASE WHEN (RTRIM(ISNULL(ADDR_LINE_2,''))<>'') THEN  RTRIM(ADDR_LINE_2) + ', ' ELSE '' END + CASE WHEN (RTRIM(ISNULL(CITY,''))<>'') THEN RTRIM(CITY) + ', ' ELSE '' END + CASE WHEN (RTRIM(ISNULL(STATE,''))<>'') THEN RTRIM(STATE) + ', ' ELSE '' END + CASE WHEN (RTRIM(ISNULL(ZIP,''))<>'') THEN RTRIM(ZIP) ELSE '' END AS AGENT_ADDRESS FROM AGENTS A LEFT OUTER JOIN WEBLOGINS W ON W.Company = A.COMPANY_CODE and W.LoginID = A.AGENT_NUMBER INNER JOIN REGION_NAMES R ON REGION_CODE = MARKETING_COMPANY WHERE ((COMPANY_CODE = '" + sCompany + "') OR (COMPANY_CODE = '07' AND START_DATE >=20091001)) and  (a.company_code='" + sCompany + "' or '" + sCompany + "' = 'ALL') and (A.status_code= '" + sStatus + "' or '" + sStatus + "' = 'ALL') ORDER BY COMPANY_CODE ASC, AGENT_NUMBER ASC");
        //In the above query for the company code -'07' the start date is given because the volume of records is more for company 07
        commPolicy.Connection = con;

        SqlDataAdapter adPolicy = new SqlDataAdapter(commPolicy.CommandText, con);
        datatab = new DataTable(); 
        adPolicy.Fill(dsPolicy);
        //adPolicy.Fill(datatab);
            if(dsPolicy != null  && dsPolicy.Tables[0] != null )
                datatab = dsPolicy.Tables[0];
        //Fill the data table for export excel
        if (datatab != null)
            dataPolicy = datatab;
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
            grdAgentList.DataSource = dt;
            grdAgentList.DataBind();
            lblcount.Text = "No Records Found for the selected criteria !!";

        }
        else
        {
            dvgrid.Style.Add("height", "600px");
            grdAgentList.DataSource = dsPolicy;
            grdAgentList.DataBind();
            if (dataPolicy != null && dataPolicy.DefaultView != null)
                dataPolicy = dataPolicy.DefaultView.ToTable();
            lblcount.Text = dsPolicy.Tables[0].Rows.Count.ToString();
        }
        }

        protected void ddlagent_PreRender(object sender, EventArgs e)
        {
            /*if (ddlHandagent.Items.FindByValue(string.Empty) == null)
            {
                ddlHandagent.Items.Insert(0, new ListItem("ALL", "ALL"));
            } */
        }

        protected void ddlregion_PreRender(object sender, EventArgs e)
        {
            /* if (ddlHandregion.Items.FindByValue(string.Empty) == null)
            {
                ddlHandregion.Items.Insert(0, new ListItem("ALL", "ALL"));
            } */
        }


        protected void ddlstate_PreRender(object sender, EventArgs e)
        {
            /*if (ddlHandstate.Items.FindByValue(string.Empty) == null)
            {
                ddlHandstate.Items.Insert(0, new ListItem("ALL", "ALL"));
            } */
        }


        protected void ddlpolicydesc_PreRender(object sender, EventArgs e)
        {
            
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            //Export("AgentLicenseList.xls", this.grdHandling);
            ExportToExcel();


        }

        protected void ExportToExcel()
        {
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                dataPolicy.DefaultView.Sort = "REGION_CODE";
                dataPolicy = dataPolicy.DefaultView.ToTable();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"Reports3P" + DateTime.Now.ToString();
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
