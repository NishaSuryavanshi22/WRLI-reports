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

namespace WRLI_Reports
{
    public partial class AgentLicenseList : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        
        DataSet dsPolicy = new DataSet();
        DataTable datatab = new DataTable(); // Create a new Data table
        public static DataTable dataPolicy = new DataTable();
        string sCompany = "15";
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
                //txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                //txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                //Session["CompanyCode"] = "15";
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                    AgentID = Session["LoginID"].ToString();

                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();
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

                SqlCommand commStates = new SqlCommand("SELECT * FROM ALLSTATES ORDER BY STATE_ABBR ASC");
                commStates.Connection = con;
                DataSet dsStates = new DataSet();
                SqlDataAdapter adStates = new SqlDataAdapter(commStates.CommandText, con);
                adStates.Fill(dsStates);
                ddlHandstate.DataSource = dsStates;
                ddlHandstate.DataTextField = "STATE_NAME";
                ddlHandstate.DataValueField = "STATE_ABBR";
                ddlHandstate.DataBind();


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

            int indexstate = ddlHandstate.SelectedItem.Value.LastIndexOf("-");
            if (indexstate > 0)
            {
                sState = ddlHandstate.SelectedItem.Value.Substring(0, indexstate);
            }

            int indexStatus = ddlHandpolicystatus.SelectedItem.Value.LastIndexOf("-");
            if (indexStatus > 0)
            {
                //sStatus = ddlHandpolicystatus.SelectedItem.Value.Substring(0, indexStatus);
            }
            if(ddlHandpolicystatus.SelectedItem.Value != null)
            {
                sStatus = ddlHandpolicystatus.SelectedItem.Value;
            }
            SqlCommand commPolicy = new SqlCommand("select A.Company_Code,A.Agent_Number,A.License_Number, W.Status_Code,CASE WHEN (NAME_FORMAT_CODE = 'C') THEN RTRIM(NAME_BUSINESS) ELSE RTRIM(INDIVIDUAL_LAST)+', '+RTRIM(INDIVIDUAL_FIRST) END AS Agent_Name,  Region_Code,CASE WHEN (ISNULL(AGENCY_NAME,'') ='') THEN Agent_Name ELSE Agency_Name END AS Region_Name, A.Licensed_State,Granted_Date,Expires_Date,License_Status,RESIDENT_IND AS Resident_State,ww.Email from AGENT_LIC_INFO A INNER JOIN AGENTS W ON W.COMPANY_CODE = A.COMPANY_CODE and W.AGENT_NUMBER = A.AGENT_NUMBER INNER JOIN REGION_NAMES R ON REGION_CODE = MARKETING_COMPANY left outer join WEBLOGINS ww on a.agent_number=ww.loginid  WHERE A.COMPANY_CODE IS NOT NULL AND ( (W.COMPANY_CODE = '16') OR (W.COMPANY_CODE = '15') OR (W.COMPANY_CODE = '05') OR (W.COMPANY_CODE = '07' AND W.START_DATE >=20091001)) and (a.company_code='" + sCompany + "' or '" + sCompany + "' = 'ALL') and (w.status_code= '" + sStatus + "' or '" + sStatus + "' = 'ALL') and (a.licensed_state= '" + sState + "' or '" + sState + "' = 'ALL') ORDER BY A.COMPANY_CODE ASC, A.AGENT_NUMBER ASC, A.LICENSED_STATE ASC");

        commPolicy.Connection = con;
        SqlDataAdapter adPolicy = new SqlDataAdapter(commPolicy.CommandText, con);
        adPolicy.Fill(dsPolicy);
        if (dsPolicy != null && dsPolicy.Tables[0] != null)
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
            if (ddlHandstate.Items.FindByValue(string.Empty) == null)
            {
                ddlHandstate.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
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
