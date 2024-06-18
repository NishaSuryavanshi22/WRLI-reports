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

namespace WRLI_Reports
{
    public partial class PerformanceReport_AllRegion : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");

        //qlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 1000");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        string agent = "WRE";
        DataSet dsPolicy = new DataSet();
        string sFromDate = DateTime.Today.ToShortDateString();
        string sToDate = DateTime.Today.ToShortDateString();
        string sCompany = "15";
        string sReportType = "NBReport_1";
        string sBussType = "BussType";
        string sCompanyName;
        DataTable datatab = new DataTable(); // Create a new Data table
        public static DataTable dataPolicy = new DataTable();

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
            if(Session["LoginID"] != null )
                agent = Session["LoginID"].ToString();
            if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                sCompany = Session["CompanyCode"].ToString();
           
            {
                Page obj;
                //Type str = obj;
                
                string stest = "";
                
                    if (Request.QueryString["Group"] != null)
                        stest = Request.QueryString["Group"].ToString();
                    lblAgent.Text = stest;
                    if (Request.QueryString["FromDate"] != null && Request.QueryString["ToDate"] != null)
                        lblRegion.Text = "Date Range: " + Request.QueryString["FromDate"].ToString() + " - " + Request.QueryString["ToDate"].ToString();
                    if (Request.QueryString["Report"] != null)
                        sBussType = Request.QueryString["Report"].ToString();

                    //lblRegion
                    //tblgrid.Visible = false;
                    //TextBox2.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                    //TextBox1.Text = System.DateTime.Now.AddMonths(-6).ToString("d");
                    if (Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "")
                        sFromDate = Request.QueryString["fromdate"].ToString();
                    if (Request.QueryString["todate"] != null && Request.QueryString["todate"] != "")
                        sToDate = Request.QueryString["todate"].ToString();
                    if (Request.QueryString["agent"] != null && Request.QueryString["agent"] != "")
                        agent = Request.QueryString["agent"].ToString();
                    if (Request.QueryString["report"] != null && Request.QueryString["report"] != "")
                        sReportType = Request.QueryString["report"].ToString();

                

                if (sBussType == "PaidReport")
                    lblReport.Text = "Paid Report";
                else
                    lblReport.Text = "New Business Report";
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


                sCompany = "15";
                //SqlCommand commStates = new SqlCommand("SELECT * FROM ALLSTATES ORDER BY STATE_ABBR ASC");
                SqlCommand commAgents = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + Session["CompanyCode"] + "' AND HIERARCHY_AGENT = '" + agent + "' ORDER BY SORTNAME ASC");
                commAgents.Connection = con;
                DataSet dsAgents = new DataSet();
                SqlDataAdapter adpAgents = new SqlDataAdapter(commAgents.CommandText, con);
                adpAgents.Fill(dsAgents);
                /*ddlListAgent.DataSource = dsAgents;
                ddlListAgent.DataTextField = "DISPLAYNAME";
                ddlListAgent.DataValueField = "AGENT_NUMBER";
                ddlListAgent.DataBind();  */

                //SqlCommand commReason = new SqlCommand("SELECT DISTINCT CONTRACT_REASON, UPPER(CONTRACT_DESC) AS CONTRACT_DESC FROM POLICIES2 WHERE RTRIM(ISNULL(CONTRACT_REASON,''))<>'' ORDER BY UPPER(CONTRACT_DESC)");
                //commReason.Connection = con;
                //DataSet dsReason = new DataSet();
                //SqlDataAdapter adReason = new SqlDataAdapter(commReason.CommandText, con);
                //adReason.Fill(dsReason);
                //ddlpolicydesc.DataSource = dsReason;
                //ddlpolicydesc.DataTextField = "CONTRACT_DESC";
                //ddlpolicydesc.DataValueField = "CONTRACT_REASON";
                //ddlpolicydesc.DataBind();

               con.Close();
               LoadGrid();
            }
        }

        protected void LoadGrid()
        {

          // string[] sFromDate = DateTime.Today.ToShortDateString();
          //  string frmDate = sFromDate[2] + sFromDate[0] + sFromDate[1];
            con.Open();

            //int index = ddlcompany.SelectedItem.Value.LastIndexOf("-");
            //if (index > 0)
            //{
            //    selectedComp = ddlcompany.SelectedItem.Value.Substring(0, index);
            //}

            //int indexagent = ddlcompany.SelectedItem.Value.LastIndexOf("-");
            //if (indexagent > 0)
            //{
            //    selectedAgent = ddlcompany.SelectedItem.Value.Substring(0, indexagent);
            //}

            //SqlCommand commPolicy = new SqlCommand("dbo.AGENT_NET_PAID_REGION_MED_EX_Test100", con);
            //SqlParameter objpara = new SqlParameter();
            //SqlParameterCollection oo;
            SqlCommand commPolicy = new SqlCommand();
            commPolicy.Connection = con;
            string sQuery = "";
            //commPolicy.CommandText = "dbo.AGENT_NET_PAID_REGION_MED_EX_Test300";
            if (sReportType == "NBReport")
                //sQuery = "Select sum(active_count)+sum(not_taken_count)+sum(suspended_count)+sum(suspended_death_count)+sum(lapse_count)+sum(death_count)+sum" + "(surrender_count)+sum(ifrescind_count)+sum(rescind_count)+sum(other_count) as total_paid,sum(sub_count) as sub,sum(active_count) as active,sum(paid_count) as paid,sum (pending_count) as pending,sum(not_taken_count) as nottaken,sum(conversion_count) as conversion,sum(declined_count) as declined,sum(declined_reapply_count) as  declined_reapply,sum(posponed_count) as posponed,sum(incomplete_count) as incomplete,sum(ineligible_count) as ineligible,sum(withdrawn_count) as withdrawn,sum(cancel_count) " + " as cancelled,sum(ifrescind_count) as ifrescind,sum(rescind_count) as rescind,sum(suspended_count) as suspended,sum(suspended_death_count) as suspended_death,sum " + " (lapse_count) as lapse,sum(death_count) as death,sum(surrender_count) as surrender,sum(other_count) as other from disposition D LEFT OUTER JOIN REGION_NAMES R ON  " + " D.REGION_CODE = R.MARKETING_COMPANY where D.agent_number in (select agent_number from agent_hierlist where COMPANY_CODE = '15' AND HIERARCHY_AGENT = 'WRE' ) and syearmonth between '201408' and '201511'  and company_code='15' AND REGION_CODE <> '' "; 
                            sQuery = " Select Region_Code as Regioncode ,dbo.GETREGIONNAME (REGION_CODE) AS Regionname,sum(sub_count) as Submitted,sum(paid_count) as paid," +
                      " sum(pending_count) as pending, sum(conversion_count) as converted,sum(declined_count) as declined,sum(declined_reapply_count) as declined_reapply," +
                      " sum(posponed_count) as posponed, sum(incomplete_count) as incomplete,sum(ineligible_count) as ineligible," +
                " sum(withdrawn_count) as withdrawn,sum(cancel_count) as cancelled  " +
                " from disposition D LEFT OUTER JOIN REGION_NAMES R ON D.REGION_CODE = R.MARKETING_COMPANY " +
                " where D.agent_number in (select agent_number from agent_hierlist where COMPANY_CODE = " + "'" + sCompany + "'"+ " AND HIERARCHY_AGENT = "+"'"+ agent + "'"+" ) " +
                " and syearmonth between '201408' and '201511'  and company_code= "+ "'" + sCompany + "'"+ " AND REGION_CODE <> '' " +
                        " group by REGION_CODE ORDER BY REGION_CODE ";

//                sQuery = " Select Region_Code as Regioncode ,dbo.GETREGIONNAME (REGION_CODE) AS Regionname,sum(sub_count) as Submitted,sum(paid_count) as paid," +
//    " sum(pending_count) as pending, sum(conversion_count) as converted,sum(declined_count) as declined,sum(declined_reapply_count) as declined_reapply," +
//    " sum(posponed_count) as posponed, sum(incomplete_count) as incomplete,sum(ineligible_count) as ineligible," +
//" sum(withdrawn_count) as withdrawn,sum(cancel_count) as cancelled  " +
//" from disposition D LEFT OUTER JOIN REGION_NAMES R ON D.REGION_CODE = R.MARKETING_COMPANY " +
//" where D.agent_number in (select agent_number from agent_hierlist where HIERARCHY_AGENT = " + "'" + agent + "'" + " ) " +
//" and syearmonth between '" + sFromDate + "' and '" + sToDate + "'  and REGION_CODE <> '' " +
//      " group by REGION_CODE ORDER BY REGION_CODE ";

            else
                //              sQuery = "Select Region_Code,dbo.GETREGIONNAME (REGION_CODE) AS Region_Name,sum(active_count)+sum(not_taken_count)+sum(suspended_count)+sum(suspended_death_count)+sum(lapse_count)+sum(death_count)+sum(surrender_count)+sum(ifrescind_count)+sum(rescind_count)+sum(other_count) as total_paid,sum(active_count) as active,sum(not_taken_count) as nottaken," +
                //" sum(rescind_count) as rescind,sum(ifrescind_count) as ifrescind,sum(suspended_count) as suspended,sum(suspended_death_count) as suspended_death," +
                //" sum(lapse_count) as lapse,sum(death_count) as death,sum(surrender_count) as surrender,sum(other_count) as other "+
                //" from disposition D LEFT OUTER JOIN REGION_NAMES R ON D.REGION_CODE = R.MARKETING_COMPANY " +
                //  " where D.agent_number in (select agent_number from agent_hierlist where COMPANY_CODE = " + "'" + sCompany + "'" + " AND HIERARCHY_AGENT = " + "'" + agent + "'" + " ) " +
                //  " and syearmonth between '201408' and '201511'  and company_code= "+ "'" + sCompany + "'" +" AND REGION_CODE <> '' "+
                //          " group by REGION_CODE ORDER BY REGION_CODE ";

                //With Respect to Comapny code.
                          sQuery = "Select Region_Code,dbo.GETREGIONNAME (REGION_CODE) AS Region_Name,sum(active_count)+sum(not_taken_count)+sum(suspended_count)+sum(suspended_death_count)+sum(lapse_count)+sum(death_count)+sum(surrender_count)+sum(ifrescind_count)+sum(rescind_count)+sum(other_count) as total_paid,sum(active_count) as active,sum(not_taken_count) as nottaken," +
                " sum(rescind_count) as rescind,sum(ifrescind_count) as ifrescind,sum(suspended_count) as suspended,sum(suspended_death_count) as suspended_death," +
                " sum(lapse_count) as lapse,sum(death_count) as death,sum(surrender_count) as surrender,sum(other_count) as other " +
                " from disposition D LEFT OUTER JOIN REGION_NAMES R ON D.REGION_CODE = R.MARKETING_COMPANY " +
                  " where D.agent_number in (select agent_number from agent_hierlist where COMPANY_CODE = " + "'" + sCompany + "'" + " AND  HIERARCHY_AGENT = " + "'" + agent + "'" + " ) " +
                              " and syearmonth between '" + sFromDate + "' and '" + sToDate + "' and company_code= " + "'" + sCompany + "'" + " AND REGION_CODE <> '' " +
                          " group by REGION_CODE ORDER BY REGION_CODE ";

  //              sQuery = "Select Region_Code,dbo.GETREGIONNAME (REGION_CODE) AS Region_Name,sum(active_count)+sum(not_taken_count)+sum(suspended_count)+sum(suspended_death_count)+sum(lapse_count)+sum(death_count)+sum(surrender_count)+sum(ifrescind_count)+sum(rescind_count)+sum(other_count) as total_paid,sum(active_count) as active,sum(not_taken_count) as nottaken," +
  //" sum(rescind_count) as rescind,sum(ifrescind_count) as ifrescind,sum(suspended_count) as suspended,sum(suspended_death_count) as suspended_death," +
  //" sum(lapse_count) as lapse,sum(death_count) as death,sum(surrender_count) as surrender,sum(other_count) as other " +
  //" from disposition D LEFT OUTER JOIN REGION_NAMES R ON D.REGION_CODE = R.MARKETING_COMPANY " +
  //  " where D.agent_number in (select agent_number from agent_hierlist where HIERARCHY_AGENT = " + "'" + agent + "'" + " ) " +
  //              " and syearmonth between '" + sFromDate + "' and '" + sToDate + "' AND REGION_CODE <> '' " +
  //          " group by REGION_CODE ORDER BY REGION_CODE ";

            commPolicy.CommandType = CommandType.StoredProcedure;
            commPolicy.CommandText = sQuery;

            //commPolicy.Parameters.AddWithValue("@company", "15");

            commPolicy.Parameters.Add("@agentid", SqlDbType.VarChar);
            commPolicy.Parameters["@agentid"].Value = agent;

            commPolicy.Parameters.Add("@company", SqlDbType.VarChar);
            commPolicy.Parameters["@company"].Value = sCompany;

            commPolicy.Parameters.Add("@fromdate", SqlDbType.VarChar);
            commPolicy.Parameters["@fromdate"].Value = "201408";

            commPolicy.Parameters.Add("@todate", SqlDbType.VarChar);
            commPolicy.Parameters["@todate"].Value = "201511";

            commPolicy.Parameters.Add("@orderby", SqlDbType.VarChar);
            commPolicy.Parameters["@orderby"].Value = "REGION_CODE";

            commPolicy.Parameters.Add("@orderdir", SqlDbType.VarChar);
            commPolicy.Parameters["@orderdir"].Value = "REGION_CODE";

            commPolicy.Parameters.Add("@resulttype", SqlDbType.VarChar);
            commPolicy.Parameters["@resulttype"].Value = "ALL";



            SqlDataAdapter adPolicy = new SqlDataAdapter(commPolicy.CommandText, con);
            adPolicy.Fill(dsPolicy);
            //commPolicy.ExecuteNonQuery();
            if (dsPolicy != null && dsPolicy.Tables[0] != null)
                datatab = dsPolicy.Tables[0];
            //Fill the data table for export excel
            if (datatab != null)
                dataPolicy = datatab;

            con.Close();

            if (dsPolicy != null && dsPolicy.Tables[0].Rows.Count == 0)
            {
                DataTable dt = new DataTable();
                DataColumn dc = new DataColumn();

                if (dt.Columns.Count == 0)
                {
                    dt.Columns.Add("REGION_CODE", typeof(string));
                    dt.Columns.Add("PAID_COUNT", typeof(string));
                    dt.Columns.Add("PAID_PREM", typeof(string));

                    dt.Columns.Add("SUB_COUNT", typeof(string));
                    dt.Columns.Add("SUB_PREM", typeof(string));
                    dt.Columns.Add("TERM_COUNT", typeof(string));
                    dt.Columns.Add("TERM_PREM", typeof(string));

                    dt.Columns.Add("NET_COUNT", typeof(string));
                    dt.Columns.Add("NET_PREM", typeof(string));

                }
                dvgrid.Style.Add("height", "120px");
                DataRow NewRow = dt.NewRow();
                dt.Rows.Add(NewRow);
                GridView1.DataSource = dt;
                GridView1.DataBind();
                lblNoRecords.Text = "No Records Found for the selected criteria !!";

            }
            else
            {
                dvgrid.Style.Add("height", "600px");
                GridView1.DataSource = dsPolicy;
                GridView1.DataBind();
                if (dataPolicy != null && dataPolicy.DefaultView != null)
                    dataPolicy = dataPolicy.DefaultView.ToTable();

                lblNoRecords.Text = dsPolicy.Tables[0].Rows.Count.ToString();
            }

        }
        protected void grdHandling_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {

                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();

       //    string Agent_num = e.Row.Cells[5].Text;
            string region_code = e.Row.Cells[0].Text;
               // string region_name = e.Row.Cells[1].Text;

                //if ((e.Row.RowType == DataControlRowType.DataRow)|| (e.Row.RowType== DataControlRowType.Header))
                //{
                //    e.Row.Cells[e.Row.Cells.Count - 1].Visible = false;
                //    GridView1.HeaderRow.Cells[e.Row.Cells.Count - 1].Visible = false;

                //}
                e.Row.Cells[0].ToolTip = "click to view details";

                string text = e.Row.Cells[0].Text;
    HyperLink link = new HyperLink();
    link.NavigateUrl = "performanceReport.aspx?Region_code=" + region_code + "&COMPANY_CODE=" + sCompany + "&Fromdate="+sFromDate+ "&Todate=" + sToDate + "&Report=" + sReportType;
                link.Text = text;
                link.Target = "_blank";
                e.Row.Cells[0].Controls.Add(link);

               //PolicyView.aspx?POLICY_NUMBER=W861860006&COMPANY_CODE=16&AGENT_NUMBER=ID01001
               // e.Row.Cells[2].Text = Convert.ToString("<a href=\"PolicyView.aspx?POLICY_NUMBER="+Policy_num+"&COMPANY_CODE="+sCompany+"&AGENT_NUMBER="+Agent_num+"Target="+"_blank"+" \">"+Policy_num+"</a>");
            }
        }


protected void Go_Click(object sender, EventArgs e)
        {

            //LoadGrid();
        }

        protected void Back_Click(object sender, EventArgs e)
        {

            //LoadGrid();

            Response.Redirect("PerformanceReport.aspx?Group=Hierarchy");
        }

        protected void ddlagent_PreRender(object sender, EventArgs e)
        {
            /*if (ddlListAgent.Items.FindByValue(string.Empty) == null)
            {
                ddlListAgent.Items.Insert(0, new ListItem("ALL", "ALL"));
            } */
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
            ExportToExcel();


        }

        protected void ExportToExcel()
        {
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                dataPolicy.DefaultView.Sort = "REGION_CODE";
                dataPolicy = dataPolicy.DefaultView.ToTable();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"Performance_AllRegion" + DateTime.Now.ToString();
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
