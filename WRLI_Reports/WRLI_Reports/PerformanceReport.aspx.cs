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
using DocumentFormat.OpenXml.VariantTypes;
using System.Configuration;
using System.Reflection;

namespace WRLI_Reports
{
    public partial class PerformanceReport : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        // SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //  SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 1000");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        string agent = "WRE";
        DataSet dsPolicy = new DataSet();
        string sCompany = string.Empty;
        bool sNB = true;
        Int32[] arr_NB = new Int32[] { };
        string ReportType = "";
        string FilterResultsType = "";
        DataTable datatab = new DataTable();

        //DataTable datatab = new DataTable(); // Create a new Data table
        public static DataTable dataPolicy = new DataTable();
        string frmDate;
        string toDate;
        string RegionCode;
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
            if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                agent = Session["LoginID"].ToString();
            

            if (!IsPostBack)
            {

                ClientScript.RegisterStartupScript(this.GetType(),"RKEY", "<script>OnPageInit('sanjeev')</script>");
                //tblgrid.Visible = false;
                RadioButton1.Checked = true;
                
                TextBox2.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                TextBox1.Text = System.DateTime.Now.AddMonths(-6).ToString("d");
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                else
                    sCompany = "15";
                //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Initial Catalog=WRE_AGENT");
                con.Open();

                //SqlCommand commCompany = new SqlCommand("select * from company_details ");
                //commCompany.Connection = con;
                //DataSet dscomp = new DataSet();
                //SqlDataAdapter adcomp = new SqlDataAdapter(commCompany.CommandText, con);
                //adcomp.Fill(dscomp);
                //List<string> lstcomp = new List<string>();
                //for (int i = 0; i < dscomp.Tables[0].Rows.Count; i++)
                //{
                //    lstcomp.Add(dscomp.Tables[0].Rows[i].ItemArray[0].ToString() + "-" + dscomp.Tables[0].Rows[i].ItemArray[1].ToString());

                //}
                //ddlcompany.DataSource = lstcomp;
                //ddlcompany.DataBind();

                //SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + agent + "' ORDER BY SORTNAME ASC");
                //comm.Connection = con;
                ////DataSet ds = new DataSet();
                //SqlDataAdapter ad = new SqlDataAdapter(comm.CommandText, con);
                //ad.Fill(dsPolicy);

                //List<string> lstagent = new List<string>();
                //for (int i = 0; i < dsPolicy.Tables[0].Rows.Count; i++)
                //{
                //    lstagent.Add(dsPolicy.Tables[0].Rows[i].ItemArray[0].ToString() + "-" + dsPolicy.Tables[0].Rows[i].ItemArray[1].ToString());

                //}
                ////Commented it is not required
                ///*ListItem obj = new ListItem();
                //obj.Value = "Region Code";
                //obj.Text = "Region Code";
                //ddlListAgent.Items.Add(obj);
                //cdStartdate.SelectedDate = DateTime.Now;
                //cdStartdate.VisibleDate = DateTime.Now; */

                ////SqlCommand commStates = new SqlCommand("SELECT * FROM ALLSTATES ORDER BY STATE_ABBR ASC");
                //SqlCommand commAgents = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany  + "' AND HIERARCHY_AGENT = '" + agent + "' ORDER BY SORTNAME ASC");
                //commAgents.Connection = con;
                //DataSet dsAgents = new DataSet();
                //SqlDataAdapter adpAgents = new SqlDataAdapter(commAgents.CommandText, con);
                //adpAgents.Fill(dsAgents);
                //ddlListAgent.DataSource = dsAgents;
                //ddlListAgent.DataTextField = "DISPLAYNAME";
                //ddlListAgent.DataValueField = "AGENT_NUMBER";
                //ddlListAgent.DataBind();

                //SqlCommand commReason = new SqlCommand("SELECT DISTINCT CONTRACT_REASON, UPPER(CONTRACT_DESC) AS CONTRACT_DESC FROM POLICIES2 WHERE RTRIM(ISNULL(CONTRACT_REASON,''))<>'' ORDER BY UPPER(CONTRACT_DESC)");
                //commReason.Connection = con;
                //DataSet dsReason = new DataSet();
                //SqlDataAdapter adReason = new SqlDataAdapter(commReason.CommandText, con);
                //adReason.Fill(dsReason);
                //ddlpolicydesc.DataSource = dsReason;
                //ddlpolicydesc.DataTextField = "CONTRACT_DESC";
                //ddlpolicydesc.DataValueField = "CONTRACT_REASON";
                //ddlpolicydesc.DataBind();

                //if (!string.IsNullOrEmpty(Request.QueryString["fromdate"]))
                //{
                //    // Redirect to the same page without any query string parameters
                //    Response.Redirect("performanceReport.aspx");
                //}

                if (Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "")
                {
                    if (Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "")
                        frmDate = Request.QueryString["fromdate"].ToString();
                    if (Request.QueryString["Todate"] != null && Request.QueryString[""] != "")
                        toDate = Request.QueryString["Todate"].ToString();
                    if (Request.QueryString["Region_code"] != null && Request.QueryString["Region_code"] != "")
                        RegionCode = Request.QueryString["Region_code"].ToString();
                    if (Request.QueryString["Company_code"] != null && Request.QueryString["Company_code"] != "")
                        sCompany = Request.QueryString["Company_code"].ToString();
                    if (Request.QueryString["Report"] != null && Request.QueryString["Report"] != "")
                        ReportType = Request.QueryString["Report"].ToString();
                    BusinessLogic();
                   // Response.Redirect("PerformanceReport.aspx");
                }



                con.Close();
               LoadDropdownValues();

                // BusinessLogic();
                hdnGridVW.Value = sNB.ToString();

              //  Response.Redirect("PerformanceReport.aspx");


            }
        }

        protected void LoadDropdownValues()
        {
            string selectedComp = "ALL";
            string selectedAgent = "ALL";

            string[] fromDate;
            string[] toDate;
            string frmDate;
            string tDate;

            if (TextBox1.Text.Contains('/'))
            {
                fromDate = TextBox1.Text.Split('/');
                 frmDate = fromDate[2] + fromDate[0] + fromDate[1];
                //
            }
            else if (TextBox1.Text.Contains('-'))
            {
                fromDate = TextBox1.Text.Split('-');
                frmDate = fromDate[2] + fromDate[0] + fromDate[1];
            }
            if (TextBox2.Text.Contains('/'))
            {
                toDate = TextBox2.Text.Split('/');
                tDate = toDate[2] + toDate[0] + toDate[1];
            }

            else if (TextBox2.Text.Contains('-'))
            {
                toDate = TextBox2.Text.Split('-');
                tDate = toDate[2] + toDate[0] + toDate[1];
            }



            //string[] fromDate = TextBox1.Text.Split('/');
            //string frmDate = fromDate[2] + fromDate[0] + fromDate[1];

            //string[] toDate = TextBox2.Text.Split('/');
            //string tDate = toDate[2] + toDate[0] + toDate[1];

            con.Open();

            SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + agent + "' ORDER BY SORTNAME ASC");
            comm.Connection = con;
            //DataSet ds = new DataSet();
            SqlDataAdapter ad = new SqlDataAdapter(comm.CommandText, con);
            ad.Fill(dsPolicy);

            List<string> lstagent = new List<string>();
            for (int i = 0; i < dsPolicy.Tables[0].Rows.Count; i++)
            {
                lstagent.Add(dsPolicy.Tables[0].Rows[i].ItemArray[0].ToString() + "-" + dsPolicy.Tables[0].Rows[i].ItemArray[1].ToString());

            }
            //Commented it is not required
            /*ListItem obj = new ListItem();
            obj.Value = "Region Code";
            obj.Text = "Region Code";
            ddlListAgent.Items.Add(obj);
            cdStartdate.SelectedDate = DateTime.Now;
            cdStartdate.VisibleDate = DateTime.Now; */

            //SqlCommand commStates = new SqlCommand("SELECT * FROM ALLSTATES ORDER BY STATE_ABBR ASC");
            SqlCommand commAgents = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + agent + "' ORDER BY SORTNAME ASC");
            commAgents.Connection = con;
            DataSet dsAgents = new DataSet();
            SqlDataAdapter adpAgents = new SqlDataAdapter(commAgents.CommandText, con);
            adpAgents.Fill(dsAgents);
            ddlListAgent.DataSource = dsAgents;
            ddlListAgent.DataTextField = "DISPLAYNAME";
            ddlListAgent.DataValueField = "AGENT_NUMBER";
            ddlListAgent.DataBind();


            //SqlCommand commreg = new SqlCommand("Select distinct d.region_code as region_code, dbo.GETREGIONNAME (REGION_CODE) as agency from disposition d where d.agent_number in (select agent_number from agent_hierlist where COMPANY_CODE = '" + sCompany + "'" + " AND HIERARCHY_AGENT = 'WRE' ) and company_code= '" +  sCompany  + "'" +" order by d.region_code  ");
            SqlCommand commreg = new SqlCommand("Select distinct d.region_code as region_code, dbo.GETREGIONNAME (REGION_CODE) as agency from disposition d where d.agent_number in (select agent_number from agent_hierlist where HIERARCHY_AGENT = 'WRE' ) order by d.region_code  ");
            commreg.Connection = con;
                DataSet dsreg = new DataSet();
                SqlDataAdapter adpreg = new SqlDataAdapter(commreg.CommandText, con);
                adpreg.Fill(dsreg);
                ddlRegion.DataSource = dsreg;
                ddlRegion.DataTextField = "region_code";
                ddlRegion.DataValueField = "region_code";
                ddlRegion.DataBind();
                ddlRegion.Items.Add("All");
                ddlRegion.Items.Add("All By Region");

                ddlBussReport.Items.Add("Paid Report");
                ddlBussReport.Items.Add("New Business Report");

        }

        protected void BusinessLogic()
        {
            //tblgrid.Visible = true;
            string[] fromDate;
            string[] toDateEntered;
            

            if ((Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "" && Request.QueryString["Todate"] != null && Request.QueryString["Todate"] != ""))
            {
                 frmDate = Request.QueryString["fromdate"].ToString();
                 toDate = Request.QueryString["Todate"].ToString();
                RegionCode = Request.QueryString["Region_code"].ToString();
                sCompany = Request.QueryString["Company_code"].ToString();
                ReportType = Request.QueryString["Report"].ToString();


            }
            else
            {
                if (TextBox1.Text.Contains('/'))
                {
                    fromDate = TextBox1.Text.Split('/');
                     frmDate = fromDate[2] + fromDate[0];
                    //
                }
                else if (TextBox1.Text.Contains('-'))
                {
                    fromDate = TextBox1.Text.Split('/');
                     frmDate = fromDate[2] + fromDate[0];
                }
                if (TextBox2.Text.Contains('/'))
                {
                    toDateEntered = TextBox2.Text.Split('/');
                    //string tDate = toDate[2] + toDate[0] + toDate[1];
                     toDate = toDateEntered[2] + toDateEntered[0];
                }

                else if (TextBox2.Text.Contains('-'))
                {
                    toDateEntered = TextBox2.Text.Split('/');
                    //string tDate = toDate[2] + toDate[0] + toDate[1];
                     toDate = toDateEntered[2] + toDateEntered[0];
                }

                //string[] fromDate = TextBox1.Text.Split('/');
                //string frmDate = fromDate[2] + fromDate[0];
                ////string frmDate = fromDate[2] + fromDate[0] + fromDate[1];

                //string[] toDateEntered = TextBox2.Text.Split('/');
                ////string tDate = toDate[2] + toDate[0] + toDate[1];
                //string toDate = toDateEntered[2] + toDateEntered[0];
                //Agent code and company code should be read from Session variable at the time of Code integration
                //string sCompany = "15";

                //  string RegionCode = "INS";

                //RegionCode = Request.QueryString["Region_code"];
                // toDate = Request.QueryString["Todate"];
                //frmDate = Request.QueryString["Fromdate"];
                //sCompany = Request.QueryString["COMPANY_CODE"];


                //if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                ReportType = ddlBussReport.SelectedItem.Text;

                RegionCode = ddlRegion.SelectedItem.Text;

                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                else
                    sCompany = "15";

                string sOrderBy = "PR.REGION_CODE";
                string sOrderDir = "";
                string sGoGreen = chkGoGreen.Checked.ToString();
                string IsGreen = "1";
                string sResultType = "ALL";
                if ((string.Compare(sGoGreen, "True") == 0))
                    IsGreen = "1";
                else
                    IsGreen = "0";
            }

            string sAgent = "WRE";

            if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                sAgent = Session["LoginID"].ToString();
            bool overview = chkCompany.Checked;


            //if (Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "")
            //    frmDate = Request.QueryString["fromdate"].ToString();
            //if (Request.QueryString["Todate"] != null && Request.QueryString["Todate"] != "")
            //    frmDate = Request.QueryString["Todate"].ToString();
            //if (Request.QueryString["Region_code"] != null && Request.QueryString["Region_code"] != "")
            //    RegionCode = Request.QueryString["Region_code"].ToString();
            //if (Request.QueryString["Company_code"] != null && Request.QueryString["Company_code"] != "")
            //    sCompany = Request.QueryString["Company_code"].ToString();

            if (con.State != ConnectionState.Open)
            //    con.Open();
            con.Open();

            SqlCommand commPolicy = new SqlCommand();
            commPolicy.Connection = con;
                //Company Overview
            commPolicy.CommandText = "Select syearmonth as Date ,sum(lapse_count) as lapse,sum(sub_count) as sub,sum(paid_count) as apaid,sum(pending_count) as pending," +
    " sum(not_taken_count) as nottaken,sum(conversion_count) as conversion,sum(declined_count) as declined,sum(declined_reapply_count) as declined_reapply, " +
    " sum(posponed_count) as posponed,sum(incomplete_count) as incomplete,sum(ineligible_count) as ineligible,sum(withdrawn_count) as withdrawn, " +
    " sum(cancel_count) as cancelled,sum(rescind_count) as rescind,sum(nbpaid_count) as paid from disposition where agent_number in " +
     " (select agent_number from agent_hierlist where (COMPANY_CODE = " + "'" + sCompany + "'" + ") AND HIERARCHY_AGENT = " + "'" + sAgent + "'" + " ) and syearmonth between " + frmDate + " and " + toDate +
     "  and company_code=" + "'" + sCompany + "'" + " group by syearmonth order by syearmonth ";
            FilterResultsType = "ALL";
            
            //NB  Hierarchy
            if (ReportType == "New Business Report")   
            {
                FilterResultsType = "1";
                if (RadioButton1.Checked && !overview) //NB  Hierarchy,NO Overview
                {
                    hdnGridVW.Value = "true";
                    sNB = true;
                    
                    commPolicy.CommandText = "Select syearmonth as Date,sum(nbpaid_count) as NBpaid,sum(pending_count) as Pending,sum(conversion_count) as Conversion," +
           " sum(declined_count) as Declined,sum(declined_reapply_count) as Declined_reapply, " +
           "sum(posponed_count) as Posponed,sum(incomplete_count) as Incomplete, sum(ineligible_count) as Ineligible, sum(withdrawn_count) as Withdrawn, " +
           " sum(cancel_count) as Cancelled,sum(paid_count) as Totalpaid from disposition where agent_number in " +
            " (select agent_number from agent_hierlist where (COMPANY_CODE = " + "'" + sCompany + "'" + ") AND HIERARCHY_AGENT = " + "'" + sAgent + "'" + " ) and syearmonth between " +
            frmDate + " and " + toDate + " and region_code= " + "'" + RegionCode + "'" + "  and company_code=" + "'" + sCompany + "'" + " group by syearmonth order by syearmonth ";
                }
                else if (RadioButton1.Checked && overview) //NB Hierarchy ,Overview
                {
                    commPolicy.CommandText = "Select syearmonth as Date,sum(nbpaid_count) as NBpaid,sum(pending_count) as Pending,sum(conversion_count) as Conversion," +
        " sum(declined_count) as Declined,sum(declined_reapply_count) as Declined_reapply, " +
        "sum(posponed_count) as Posponed,sum(incomplete_count) as Incomplete, sum(ineligible_count) as Ineligible, sum(withdrawn_count) as Withdrawn, " +
        " sum(cancel_count) as Cancelled,sum(paid_count) as Totalpaid from disposition where agent_number in " +
         " (select agent_number from agent_hierlist where (COMPANY_CODE = " + "'" + sCompany + "'" + ") AND HIERARCHY_AGENT = " + "'" + sAgent + "'" + " ) and syearmonth between " + frmDate + " and " + toDate +
        " and company_code=" + "'" + sCompany + "'" + " group by syearmonth order by syearmonth ";
                }


                else if (RadioButton2.Checked && !overview) //NB Individual ,NO Overview
                {
                    
                    hdnGridVW.Value = "true";
                    sNB = true;
                    commPolicy.CommandText = "Select syearmonth as Date,sum(nbpaid_count) as NBpaid,sum(pending_count) as Pending,sum(conversion_count) as Conversion," +
          " sum(declined_count) as Declined,sum(declined_reapply_count) as Declined_reapply, " +
          "sum(posponed_count) as Posponed,sum(incomplete_count) as Incomplete, sum(ineligible_count) as Ineligible, sum(withdrawn_count) as Withdrawn, " +
          " sum(cancel_count) as Cancelled,sum(paid_count) as Totalpaid from disposition where agent_number = " + "'" + sAgent + "'" + " and syearmonth between " +
           frmDate + " and " + toDate + " and region_code= " + "'" + RegionCode + "'" + "  and company_code=" + "'" + sCompany + "'" + " group by syearmonth order by syearmonth ";
                }
                else if (RadioButton2.Checked && overview) //NB Individual, Overview
                {
                    hdnGridVW.Value = "false";
                    sNB = false;
                    //FilterResultsType = "3";
                    commPolicy.CommandText = "Select syearmonth as Date,sum(nbpaid_count) as NBpaid,sum(pending_count) as Pending,sum(conversion_count) as Conversion," +
          " sum(declined_count) as Declined,sum(declined_reapply_count) as Declined_reapply, " +
          "sum(posponed_count) as Posponed,sum(incomplete_count) as Incomplete, sum(ineligible_count) as Ineligible, sum(withdrawn_count) as Withdrawn, " +
          " sum(cancel_count) as Cancelled,sum(paid_count) as Totalpaid from disposition where agent_number = " +  "'" + sAgent + "'" + " and syearmonth between " +
           frmDate + " and " + toDate + "  and company_code=" + "'" + sCompany + "'" + " group by syearmonth order by syearmonth ";

                }
            }
            else if (ReportType == "Paid Report" )   //Paid Report 
            {
                FilterResultsType = "2";
                hdnGridVW.Value = "false";
                sNB = false;
                if (RadioButton1.Checked && !overview) //Paid Report Hier, NO Overview
                {
                    commPolicy.CommandText = "Select syearmonth as Date,sum(sub_count) as Sub,sum(active_count) as Active,sum(rescind_count) as Rescinded,sum(not_taken_count) as Nottaken," +
    " sum(paid_count) as Paid,sum(suspended_count) as Suspended,sum(suspended_death_count) as Suspended_death,sum(lapse_count) as Lapse," +
    " sum(death_count) as Death,sum(surrender_count) as Surrender,sum(other_count) as Other,sum(ifrescind_count) as InforceRescissions from disposition " +
      " where agent_number in  (select agent_number from agent_hierlist where (COMPANY_CODE = " + "'" + sCompany + "'" + ") AND HIERARCHY_AGENT = " + "'" + sAgent + "'" + " )" +
      " and syearmonth between " + frmDate + " and " + toDate + " and region_code= " + "'" + RegionCode + "'" + "  and company_code=" + "'" + sCompany + "'" + " group by syearmonth order by syearmonth ";
                }
                else if (RadioButton1.Checked && overview) //Paid Report Hier, Overview
                {
                    commPolicy.CommandText = "Select syearmonth as Date,sum(sub_count) as Sub,sum(active_count) as Active,sum(rescind_count) as Rescinded,sum(not_taken_count) as Nottaken," +
    " sum(paid_count) as Paid,sum(suspended_count) as Suspended,sum(suspended_death_count) as Suspended_death,sum(lapse_count) as Lapse," +
    " sum(death_count) as Death,sum(surrender_count) as Surrender,sum(other_count) as Other,sum(ifrescind_count) as InforceRescissions from disposition " +
      " where agent_number  in  (select agent_number from agent_hierlist where (COMPANY_CODE = " + "'" + sCompany + "'" + ") AND HIERARCHY_AGENT = " + "'" + sAgent + "'" + " )" +
      " and syearmonth between " + frmDate +  " and " + toDate + "  and company_code=" + "'" + sCompany + "'" + " group by syearmonth order by syearmonth ";
                }

                else if (RadioButton2.Checked && !overview) //Paid Report Indi, No Overview
                {
                    commPolicy.CommandText = "Select syearmonth as Date,sum(sub_count) as Sub,sum(active_count) as Active,sum(rescind_count) as Rescinded,sum(not_taken_count) as Nottaken," +
    " sum(paid_count) as paid,sum(suspended_count) as suspended,sum(suspended_death_count) as suspended_death,sum(lapse_count) as lapse," +
    " sum(death_count) as death,sum(surrender_count) as surrender,sum(other_count) as other,sum(ifrescind_count) as InforceRescissions from disposition " +
      " where agent_number= " + "'" + sAgent + "'" + " and syearmonth between " + frmDate +
        " and " + toDate + " and region_code= " + "'" + RegionCode + "'" + "  and company_code=" + "'" + sCompany + "'" + " group by syearmonth order by syearmonth ";
                }
                else if (RadioButton2.Checked && overview) //Paid Report Indi, Overview
                {
                    commPolicy.CommandText = "Select syearmonth as Date,sum(sub_count) as sub,sum(active_count) as active,sum(rescind_count) as rescinded,sum(not_taken_count) as nottaken," +
   " sum(paid_count) as paid,sum(suspended_count) as suspended,sum(suspended_death_count) as suspended_death,sum(lapse_count) as lapse," +
   " sum(death_count) as death,sum(surrender_count) as surrender,sum(other_count) as other,sum(ifrescind_count) as InforceRescissions from disposition " +
     " where agent_number= " + "'" + sAgent + "'" + " and syearmonth between " + frmDate +
       " and " + toDate  + "  and company_code=" + "'" + sCompany + "'" + " group by syearmonth order by syearmonth ";
                }
            }
                
            commPolicy.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter adPolicy = new SqlDataAdapter(commPolicy.CommandText, con);
            adPolicy.Fill(dsPolicy);
            //commPolicy.ExecuteNonQuery();
            con.Close();

            if (dsPolicy != null && dsPolicy.Tables[0].Rows.Count == 0)
            {
                DataTable dt = new DataTable();
                DataColumn dc = new DataColumn();

                if (dt.Columns.Count == 0)
                {
                    lblcount.Text = "No Records Found for the selected criteria !!!";
                    dvgrid.Visible = false;
                    //lblcount.Visible = false;
                    /*if (FilterResultsType == "1")
                    {
                        dt.Columns.Add("Active", typeof(string));
                        dt.Columns.Add("Not taken", typeof(string));
                        dt.Columns.Add("Rescinded", typeof(string));

                        dt.Columns.Add("InforceRescissions", typeof(string));
                        dt.Columns.Add("suspended", typeof(string));
                        dt.Columns.Add("suspended_death", typeof(string));
                        dt.Columns.Add("death", typeof(string));

                        dt.Columns.Add("lapse", typeof(string));
                        dt.Columns.Add("surrender", typeof(string));
                    }
                    else if (FilterResultsType == "2")
                    {

                        lblcount.Text = "No Records Found for the selected criteria !!";
                    }*/

                }
                dvgrid.Style.Add("height", "120px");
                DataRow NewRow = dt.NewRow();
                dt.Rows.Add(NewRow);
                GridView1.DataSource = dt;
                GridView1.DataBind();

            }
            else
            {
                dvgrid.Visible = true;
                lblcount.Visible = true;
                AddEditRows();
                //NBWrapper objNB = BindResults(dsPolicy);
                dvgrid.Style.Add("height", "600px");
                if (dataPolicy != null && dataPolicy.DefaultView != null)
                    dataPolicy = dataPolicy.DefaultView.ToTable();
                //GridView1.DataSource = dsPolicy;
                //GridView1.DataBind();

              //  Response.Redirect("PerformanceReport.aspx");


            }
            if (dsPolicy.Tables[0].Rows.Count == 0)
                 lblcount.Text = "No Records Found for the selected criteria !!!";
            else
                lblcount.Text = "Total Policy Count: " + dsPolicy.Tables[0].Rows.Count.ToString();

            PropertyInfo isreadonly = typeof(System.Collections.Specialized.NameValueCollection).GetProperty("IsReadOnly", BindingFlags.Instance | BindingFlags.NonPublic);
            // make collection editable
            isreadonly.SetValue(this.Request.QueryString, false, null);
            // remove
            this.Request.QueryString.Remove("fromdate");
        }


        protected void InitPaidReportColumns(int Rowcount)
        {
            arr_NB[0] = new Int32();
            arr_NB[1] = new Int32();
            arr_NB[2] = new Int32();
            arr_NB[3] = new Int32();
            arr_NB[4] = new Int32();
            arr_NB[5] = new Int32();
            arr_NB[6] = new Int32();
            arr_NB[7] = new Int32();
            arr_NB[8] = new Int32();
            arr_NB[9] = new Int32();
            arr_NB[10] = new Int32();
            arr_NB[11] = new Int32();
            //arr_NB[12] = new Int32();
            //arr_NB[13] = new Int32();
            

        }

        protected void InitNBColumns(int Rowcount)
        {
            arr_NB[0] = new Int32();
            arr_NB[1] = new Int32();
            arr_NB[2] = new Int32();
            arr_NB[3] = new Int32();
            arr_NB[4] = new Int32();
            arr_NB[5] = new Int32();
            arr_NB[6] = new Int32();
            arr_NB[7] = new Int32();
            arr_NB[8] = new Int32();
            arr_NB[9] = new Int32();
            arr_NB[10] = new Int32();
            arr_NB[11] = new Int32();
            //arr_NB[12] = new Int32();
            //arr_NB[13] = new Int32();

            //arr_NB[14] = new Int32();
            //arr_NB[15] = new Int32();
            //arr_NB[16] = new Int32();

        }

        protected void PRReadColumnValues(DataRowView DataRowCurrView,ref Int32[] arr_NB)
        {
            arr_NB[1] += Convert.ToInt32(DataRowCurrView["sub"]);
            arr_NB[2] += Convert.ToInt32(DataRowCurrView["active"]);
            arr_NB[3] += Convert.ToInt32(DataRowCurrView["nottaken"]);
            arr_NB[4] += Convert.ToInt32(DataRowCurrView["rescinded"]);

            arr_NB[5] += Convert.ToInt32(DataRowCurrView["InforceRescissions"]);
            arr_NB[6] += Convert.ToInt32(DataRowCurrView["suspended"]);

            arr_NB[7] += Convert.ToInt32(DataRowCurrView["suspended_death"]);
            arr_NB[8] += Convert.ToInt32(DataRowCurrView["death"]);

            arr_NB[9] += Convert.ToInt32(DataRowCurrView["lapse"]);
            arr_NB[10] += Convert.ToInt32(DataRowCurrView["surrender"]);
            arr_NB[11] += Convert.ToInt32(DataRowCurrView["other"]);
            arr_NB[12] += Convert.ToInt32(DataRowCurrView["paid"]);

            //MyDataView1.RowFilter = "ProductName LIKE 'A%'";
            //DataView MyDataView2 = new DataView(datatbl);
            //return DataRowCurrView;

        }

        protected void NBReadColumnValues(DataRowView DataRowCurrView, ref Int32[] arr_NB)
        {
            arr_NB[0] += Convert.ToInt32(DataRowCurrView["Date"]);
            arr_NB[1] += Convert.ToInt32(DataRowCurrView["NBpaid"]);
            arr_NB[2] += Convert.ToInt32(DataRowCurrView["pending"]);
            arr_NB[3] += Convert.ToInt32(DataRowCurrView["conversion"]);
            arr_NB[4] += Convert.ToInt32(DataRowCurrView["declined"]);
            arr_NB[5] += Convert.ToInt32(DataRowCurrView["declined_reapply"]);
            arr_NB[6] += Convert.ToInt32(DataRowCurrView["posponed"]);
            
            arr_NB[7] += Convert.ToInt32(DataRowCurrView["incomplete"]);
            arr_NB[8] += Convert.ToInt32(DataRowCurrView["ineligible"]);
            arr_NB[9] += Convert.ToInt32(DataRowCurrView["withdrawn"]);
            arr_NB[10] += Convert.ToInt32(DataRowCurrView["cancelled"]);
            arr_NB[11] += Convert.ToInt32(DataRowCurrView["Totalpaid"]);
           // arr_NB[15] += Convert.ToInt32(DataRowCurrView["rescind"]);
           // arr_NB[5] += Convert.ToInt32(DataRowCurrView["nottaken"]);
            
            //arr_NB[1] += Convert.ToInt32(DataRowCurrView["lapse"]);
            //arr_NB[2] += Convert.ToInt32(DataRowCurrView["sub"]);
            
            
        }

        protected void AddEditRows()
        {

            DataTable datatbl = dsPolicy.Tables[0];
            DataView MyDataView1 = new DataView(datatbl);
            DataRowView DataRowCurrView = null;
            MyDataView1.AllowNew = true;

            
            //MyDataRowView["active"] = 111;
            //MyDataRowView["sub"] = 222;
           
            
            int nRowct = datatbl.Rows.Count;
            int nColCt = datatbl.Columns.Count;
            arr_NB = new Int32[nColCt];
            if (FilterResultsType == "1")
                InitNBColumns(nColCt);
            else if (FilterResultsType == "2")
                InitPaidReportColumns(nColCt);
            else if (FilterResultsType == "ALL")
            {
                //InitNBColumns(nColCt);
            }
            //
            for (int nIndex = 0; nIndex < nRowct; nIndex++)
            {

                DataRowCurrView = MyDataView1[nIndex];
                //List<Int32> stringList = new List<Int32>();

                if (FilterResultsType == "1" )
                {
                    NBReadColumnValues(DataRowCurrView, ref arr_NB);
                }
                else if (FilterResultsType == "2" )
                {
                    PRReadColumnValues(DataRowCurrView, ref arr_NB);
                    
                }
                else if (FilterResultsType == "ALL")
                {
                }
                // End of new logic
            }
                //Below line Commented by Siva 07 Jan
                DataRowView MyDataRowView = MyDataView1.AddNew();
                //DataRow MyDataRowView = MyDataView1.Table.NewRow();
                //MyDataView1.Table.Rows.InsertAt(MyDataRowView, 0); 

                int position = 0;
                int i = 0;
                MyDataView1.AllowEdit = true;
                MyDataRowView.BeginEdit();

                if (FilterResultsType == "1" )
                {

                    MyDataRowView["Date"] = "Total";
                    MyDataRowView["NBPaid"] = arr_NB[1];
                    MyDataRowView["Pending"] = arr_NB[2];
                    MyDataRowView["conversion"] = arr_NB[3];
                    MyDataRowView["declined"] = arr_NB[4];
                    MyDataRowView["declined_reapply"] = arr_NB[5];
                    MyDataRowView["posponed"] = arr_NB[6];
                    MyDataRowView["incomplete"] = arr_NB[7];
                    MyDataRowView["ineligible"] = arr_NB[8];
                    MyDataRowView["withdrawn"] = arr_NB[9];
                    MyDataRowView["cancelled"] = arr_NB[10];
                    MyDataRowView["Totalpaid"] = arr_NB[11];
                    DataRowCurrView.EndEdit();
              }
               else if (FilterResultsType == "2")
               {

                    MyDataRowView["Date"] = "Total";
                    MyDataRowView["sub"] = arr_NB[1];
                    MyDataRowView["active"] = arr_NB[2];
                    MyDataRowView["nottaken"] = arr_NB[3];
                    MyDataRowView["rescinded"] = arr_NB[4];
                    MyDataRowView["InforceRescissions"] = arr_NB[5];
                    //MyDataRowView["nottaken"] = arr_NB[5];
                    MyDataRowView["suspended"] = arr_NB[6];
                    MyDataRowView["suspended_death"] = arr_NB[7];
                    MyDataRowView["death"] = arr_NB[8];
                    MyDataRowView["lapse"] = arr_NB[9];
                    MyDataRowView["surrender"] = arr_NB[10];
                    MyDataRowView["other"] = arr_NB[11];
                   MyDataRowView["paid"] = arr_NB[12];

            }

            //Start New logic to insert
            position = i + 1; //Dont want to insert at the row, but after.
            /*DataRow newRow = MyDataView1.Table.NewRow();
            newRow[0] = "Net Total";
            newRow[1] = arr_NB[1];
            MyDataView1.Table.Rows.InsertAt(newRow, 0); */
            //End
            //Data binding
            GridView1.DataSource = MyDataView1;
            GridView1.DataBind();
            //datatab = MyDataView1.Table.Clone();
            datatab = MyDataView1.Table.Copy();


        }

        protected void Go_Click(object sender, EventArgs e)
        {

            //string[] sFmDate;
            //string[] sTDate;
            //if (TextBox1.Text.Contains('/'))
            //{
            //    sFmDate = TextBox1.Text.Split('/');
            //    string sFromDate = sFmDate[2] + sFmDate[0];
             
            //}
            //else if (TextBox1.Text.Contains('-'))
            //{
            //    sFmDate = TextBox1.Text.Split('/');
            //    string sFromDate = sFmDate[2] + sFmDate[0];
            //}
            //if (TextBox2.Text.Contains('/'))
            //{
            //     sTDate = TextBox2.Text.Split('/');
            //    string sToDate = sTDate[2] + sTDate[0];
            //}

            //else if (TextBox2.Text.Contains('-'))
            //{
            //     sTDate = TextBox2.Text.Split('/');
            //    string sToDate = sTDate[2] + sTDate[0];
            //}

            // string sFromDate = TextBox1.Text;
          string[] sFmDate = TextBox1.Text.Split('/');
          string sFromDate = sFmDate[2] + sFmDate[0];

           //// string sToDate = TextBox2.Text;
            string[] sTDate = TextBox2.Text.Split('/');
            string sToDate = sTDate[2] + sTDate[0];

            string PaidReport;
            if (ddlBussReport.SelectedItem.Text != "Paid Report")
            {
                ReportType = "NBReport";
                //bool IsHierarchy = true;
            }
            else
            {
                ReportType = "Paid Report";

            }
            if (ddlRegion.SelectedItem.Text == "All By Region")
            {
                
                string sHeir = "Heirarachy is Heirarachy";
                //bool IsHierarchy = true;
                if (!RadioButton1.Checked)
                sHeir = "Heirarchy is Individual";
                Response.Redirect("PerformanceReport_AllRegion.aspx?Group=" + sHeir + "&FromDate=" + sFromDate + "&ToDate=" + sToDate + "&Report=" + ReportType);
            }
            else if (ddlBussReport.SelectedItem.Text == "All")
                BusinessLogic();
            else
                BusinessLogic();
                
        }

        protected void ddlagent_PreRender(object sender, EventArgs e)
        {
            if (ddlListAgent.Items.FindByValue(string.Empty) == null)
            {
                ddlListAgent.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
        }

        protected void ddlSort_PreRender(object sender, EventArgs e)
        {
            //if (ddlregion.Items.FindByValue(string.Empty) == null)
            //{
            //    ddlregion.Items.Insert(0, new ListItem("ALL", "ALL"));
            //}
        }



        protected void DiableIndividualAgent(object sender, EventArgs e)
        {
            if (chkCompany.Checked)
                ddlRegion.Enabled = false;
            else
                ddlRegion.Enabled = true;

        }

        protected void OnClientDateSelectionChanged(object sender, EventArgs e)
        {
        }
        protected void ddlBussReport_PreRender(object sender, EventArgs e)
        {

        }
        protected void ddlRegion_PreRender(object sender, EventArgs e)
        {
        }

        protected void ddlstate_PreRender(object sender, EventArgs e)
        {
            //if (ddlstate.Items.FindByValue(string.Empty) == null)
            //{
            //    ddlstate.Items.Insert(0, new ListItem("ALL", "ALL"));
            //}
        }


        protected void ddlpolicydesc_PreRender(object sender, EventArgs e)
        {
            //if (ddlpolicydesc.Items.FindByValue(string.Empty) == null)
            //{
            //    ddlpolicydesc.Items.Insert(0, new ListItem("ALL", "ALL"));
            //}
        }

        protected void Button2_Click(object sender, EventArgs e)
        {

            //Export("PolicyReport.xls", this.GridView1);
            BusinessLogic();
            ExportToExcel();
        }

        protected void ExportToExcel()
        {
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                dataPolicy.DefaultView.Sort = "REGION_CODE";
                dataPolicy = dataPolicy.DefaultView.ToTable();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"PerformanceReport" + DateTime.Now.ToString();
                Response.AddHeader("Content-Disposition", "inline;filename=" + filename.Replace("/", "").Replace(":", "") + ".xlsx");
                //Call  Export function
                //Response.BinaryWrite(ExportToCSVFileOpenXML(datatab));   
                Response.BinaryWrite(Utils.ExportToCSVFileOpenXML(dataPolicy));
                Response.Flush();
                Response.End();
            }
        }


        //public static DataTable CreateTable(DataView obDataView)
        //{
        //    if (null == obDataView)
        //    {
        //        throw new ArgumentNullException
        //        ("DataView", "Invalid DataView object specified");
        //    }

        //    DataTable obNewDt = obDataView.Table.Clone();
        //    int idx = 0;
        //    string[] strColNames = new string[obNewDt.Columns.Count];
        //    foreach (DataColumn col in obNewDt.Columns)
        //    {
        //        strColNames[idx++] = col.ColumnName;
        //    }
        //    IEnumerator viewEnumerator = obDataView.GetEnumerator();
        //    while (viewEnumerator.MoveNext())
        //    {
        //        DataRowView drv = (DataRowView)viewEnumerator.Current;
        //        DataRow dr = obNewDt.NewRow();
        //        try
        //        {
        //            foreach (string strName in strColNames)
        //            {
        //                dr[strName] = drv[strName];
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            //  Trace.WriteLine(ex.Message);
        //        }
        //        obNewDt.Rows.Add(dr);
        //    }
        //    return obNewDt;
        //}

       

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

        

        public NBWrapper BindResults(DataSet objNB)
        {
            NewBusiness objNewBuss = new NewBusiness() { };
            NBWrapper objContainer = new NBWrapper();
            
            if (objNB.Tables != null)
            {
                object[]  objj;
                int nRowLength = objNB.Tables[0].Rows.Count;
                
                int nRowCt = objNB.Tables.Count;
               // objContainer.objWrapper[] = new NewBusiness[]{};

                for (int nIndex = 0; nIndex < nRowLength; nIndex++)
                {
                    int nColumnCt = objNB.Tables[0].Rows[nIndex].ItemArray.Length;
                    objj = objNB.Tables[0].Rows[nIndex].ItemArray;
                    for (int nInd = 0; nInd < nColumnCt -1; nInd++)
                    {
                        //objContainer.objWrapper[nRowLength] = new NewBusiness[nRowLength];
                        objContainer.objWrapper = new NewBusiness[nRowLength];
                        objContainer.objWrapper[nIndex] = new NewBusiness();
                        objContainer.objWrapper[nIndex].Lapse = objNB.Tables[0].Rows[nIndex].ItemArray[nInd].ToString();
                        
                        //objNewBuss = objNB.Tables[nIndex].Rows[nInd].ToString();
                    }

                }
                return objContainer;
            }
            return objContainer ;
        }
       
    }
    public class NewBusiness
    {
        public string SubmittedDate;
        public string Lapse;
        public string Submitted;
        public string Paid;
        public string Pending;
        public string Nottaken;
        public string Conversion;
        public string Declined;
        public string Declined_Reapply;
        public string Postponed;
        public string Incomplete;
        public string Ineligible;
        public string WithDrawn;
        public string Cancelled;
        string Rescind;
        string NBPaidCount;
        string Total;
    }
    public class NBWrapper
    {
        public NewBusiness[] objWrapper = new NewBusiness[] { };
    }

}
