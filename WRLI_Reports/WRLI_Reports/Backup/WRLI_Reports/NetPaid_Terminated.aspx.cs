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
    public partial class _NetPaidTerminated : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());

        SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        string HierAgent = "WRE";
        string agent ="WRE";
        DataSet dsPolicy = new DataSet();

        DataTable datatab = new DataTable(); // Create a new Data table
        public static DataTable dataPolicy = new DataTable();
        //
        string sCompany = "15";
        string sRegionCode = "INS";
        string AgentID = "WRE";
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
                HierAgent = Session["LoginID"].ToString();
            //agent = Session["LoginID"].ToString();
            if (!IsPostBack)
            {
                tblgrid.Visible = false;
                TextBox2.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                TextBox1.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                Session["CompanyCode"] = "15";
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

                SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + Session["CompanyCode"] + "' AND HIERARCHY_AGENT = '" + AgentID + "' ORDER BY SORTNAME ASC");
                comm.Connection = con;
                //DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(comm.CommandText, con);
                ad.Fill(dsPolicy);
                //
                if (dsPolicy != null && dsPolicy.Tables[0] != null)
                    datatab = dsPolicy.Tables[0];

                List<string> lstagent = new List<string>();
                for (int i = 0; i < dsPolicy.Tables[0].Rows.Count; i++)
                {
                    lstagent.Add(dsPolicy.Tables[0].Rows[i].ItemArray[0].ToString() + "-" + dsPolicy.Tables[0].Rows[i].ItemArray[1].ToString());

                }
                //Commented it is not required
                //ddlagent.DataSource = lstagent;
                //ddlagent.DataBind();
                ListItem obj = new ListItem();
                obj.Value = "Region Code";
                obj.Text = "Region Code";
                ddlagent.Items.Add(obj);

                //SqlCommand commStates = new SqlCommand("SELECT * FROM ALLSTATES ORDER BY STATE_ABBR ASC");
                //commStates.Connection = con;
                //DataSet dsStates = new DataSet();
                //SqlDataAdapter adStates = new SqlDataAdapter(commStates.CommandText, con);
                //adStates.Fill(dsStates);
                //ddlstate.DataSource = dsStates;
                //ddlstate.DataTextField = "STATE_NAME";
                //ddlstate.DataValueField = "STATE_ABBR";
                //ddlstate.DataBind();

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
            }
        }

        protected void InvokeSP()
        {
            string[] fromDate = TextBox1.Text.Split('/');
            string frmDate = fromDate[2] + fromDate[0] + fromDate[1];

            string[] toDate = TextBox2.Text.Split('/');
            string tDate = toDate[2] + toDate[0] + toDate[1];
            if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                HierAgent = Session["LoginID"].ToString();
            //string sAgent = Session["LoginID"].ToString();
            if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                sCompany = Session["CompanyCode"].ToString();
            string sOrderBy = "PR.REGION_CODE";
            string sOrderDir = "ASC";
            string sGoGreen = chkGoGreen.Checked.ToString();
            string IsGreen = "1";
            string sResultType = "ALL";
            if ((string.Compare(sGoGreen, "True") == 0))
                IsGreen = "1";
            else
                IsGreen = "0";
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

            SqlCommand commPolicy = new SqlCommand();
            commPolicy.Connection = con;
            commPolicy.CommandType = CommandType.StoredProcedure;
            //commPolicy.CommandText = "dbo.AGENT_NET_PAID_REGION_MED_EX";


            commPolicy.Parameters.AddWithValue("@agentid", SqlDbType.VarChar).Value = HierAgent;
            commPolicy.Parameters.AddWithValue("@company", SqlDbType.VarChar).Value = sCompany;
            commPolicy.Parameters.AddWithValue("@fromdate", SqlDbType.VarChar).Value = frmDate;
            commPolicy.Parameters.AddWithValue("@todate", SqlDbType.VarChar).Value = tDate;
            commPolicy.Parameters.AddWithValue("@orderby", SqlDbType.VarChar).Value = sOrderBy;
            commPolicy.Parameters.AddWithValue("@orderdir", SqlDbType.VarChar).Value = sOrderDir;
            commPolicy.Parameters.AddWithValue("@resulttype", SqlDbType.VarChar).Value = sResultType;
            commPolicy.Parameters.AddWithValue("@green", SqlDbType.VarChar).Value = IsGreen;

            SqlDataAdapter adPolicy = new SqlDataAdapter(commPolicy);
            commPolicy.CommandText = "dbo.AGENT_NET_PAID_REGION_MED_EX";

            adPolicy.Fill(dsPolicy);
            if (dsPolicy != null && dsPolicy.Tables[0] != null)
                datatab = dsPolicy.Tables[0];

            //commPolicy.ExecuteNonQuery();
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
                lblcount.Text = "No Records Found for the selected criteria !!";

            }
            else
            {
                dvgrid.Style.Add("height", "600px");
                GridView1.DataSource = dsPolicy;
                GridView1.DataBind();
                if (dataPolicy != null && dataPolicy.DefaultView != null)
                    dataPolicy = dataPolicy.DefaultView.ToTable();

                if (dsPolicy.Tables[0] != null)
                {
                    lblcount.Text = dsPolicy.Tables[0].Rows.Count.ToString();
                    datatab = dsPolicy.Tables[0];
                }
            }
        }




        protected void Button1_Click(object sender, EventArgs e)
        {
            tblgrid.Visible = true;
            string selectedComp = "ALL";
            string selectedAgent = "ALL";
            InvokeSP();
        }

        protected void ddlagent_PreRender(object sender, EventArgs e)
        {
            if (ddlagent.Items.FindByValue(string.Empty) == null)
            {
                ddlagent.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
        }

        protected void ddlSort_PreRender(object sender, EventArgs e)
        {
            //if (ddlregion.Items.FindByValue(string.Empty) == null)
            //{
            //    ddlregion.Items.Insert(0, new ListItem("ALL", "ALL"));
            //}
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
            //Export("NetPaid_Terminated.xls", this.GridView1);
            ExportToExcel();
        }

        protected void ExportToExcel()
        {
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                dataPolicy.DefaultView.Sort = "REGION_CODE";
                dataPolicy = dataPolicy.DefaultView.ToTable();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"NetPaid_Terminated" + DateTime.Now.ToString();
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
