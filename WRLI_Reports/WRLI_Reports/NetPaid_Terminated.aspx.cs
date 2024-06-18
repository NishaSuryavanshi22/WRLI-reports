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
    public partial class _NetPaidTerminated : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        // SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 3600");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        string HierAgent = "WRE";
        String sOrder;
        string agent ="WRE";
        DataSet dsPolicy = new DataSet();

        DataTable datatab = new DataTable(); // Create a new Data table
        public static DataTable dataPolicy = new DataTable();
        //
        string sCompany = "15";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        private string frmDate;
        private string tDate;
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

            string[] fromDate;
            string[] toDate;
            if (TextBox1.Text.Contains('/'))
            {
                fromDate = TextBox1.Text.Split('/');
                frmDate = fromDate[2] + fromDate[0] + fromDate[1];
                
            }
            else if (TextBox1.Text.Contains('-'))
            {
                fromDate = TextBox1.Text.Split('-');
                string frmDate = fromDate[2] + fromDate[0] + fromDate[1];
            }
            if (TextBox2.Text.Contains('/'))
            {
                toDate = TextBox2.Text.Split('/');
                tDate = toDate[2] + toDate[0] + toDate[1];
            }

            else if (TextBox2.Text.Contains('-'))
            {
                toDate = TextBox2.Text.Split('-');
                string tDate = toDate[2] + toDate[0] + toDate[1];
            }


            //string[] fromDate = TextBox1.Text.Split('/');
            //string frmDate = fromDate[2] + fromDate[0] + fromDate[1];

            //string[] toDate = TextBox2.Text.Split('/');
            //string tDate = toDate[2] + toDate[0] + toDate[1];


            if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                HierAgent = Session["LoginID"].ToString();
            //string sAgent = Session["LoginID"].ToString();
            if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                sCompany = Session["CompanyCode"].ToString();
            string sOrderBy = "PR.REGION_CODE";
            string sOrderDir;
            string sGoGreen = chkGoGreen.Checked.ToString();
            string IsGreen = "1";
            string sResultType = "ALL";
            if ((string.Compare(sGoGreen, "True") == 0))
                IsGreen = "1";
            else
                IsGreen = "0";

            string bType = "ALL";
            if (rdListType.Items[0].Selected)
            {
                bType = "ALL";
            }
            else    if (rdListType.Items[2].Selected)
            {
                bType = "NON-MED";
            }
            else if (rdListType.Items[1].Selected)
            {
                bType = "MED";
            }

            if (ddlSort.SelectedItem.Value != null)
            {
               sOrder = ddlSort.SelectedItem.Value;
            }
            if (sOrder == "Descending")
            {
                sOrderDir= "DESC";
            }
            else
            {
                sOrderDir = "ASC";
            }
            

            

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
            commPolicy.CommandTimeout = 0;
          //  commPolicy.CommandTimeout = 9000;

            commPolicy.CommandType = CommandType.StoredProcedure;
            commPolicy.CommandText = "dbo.AGENT_NET_PAID_REGION_MED_EX";
            SqlDataAdapter adPolicy = new SqlDataAdapter(commPolicy.CommandText, con);

            //            commPolicy.CommandText = "select(UR.REGION_CODE),"+
            // "sum(TERM_COUNT) as TERM_COUNT," +
            // "SUM(NET_COUNT) as NET_COUNT, " + 
            //                "  sum(PAID_COUNT) as PAID_COUNT" +
            // " from NETREGION UR inner join ]]== PO ON  PO.AGENT_NUMBER = UR.AGENT_NUMBER left outer join U_REQ PB ON PB.AGENT_NUMBER = UR.AGENT_NUMBER" +
            //" WHERE UR.AGENT_NUMBER in " +
            //" (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = 'WRE')  AND " +
            // "PB.COMPANY_CODE = '07' AND UR.REGION_CODE IS NOT NULL AND" +
            // "(UR.PAY_DATE BETWEEN '"+frmDate+"' AND '"+tDate+"') AND" +
            // "((LEFT(UR.REGION_CODE, 1) = 'ALL') OR(SUBSTRING(UR.REGION_CODE, 2, 1) = 'ALL') OR(LEFT(UR.REGION_CODE, 2) = 'ALL') OR('ALL' = 'ALL')) AND((MED_TYPE = '" + bType + "') OR " +
            //            " ('" + bType + "' = 'ALL') OR" +
            // "(MED_TYPE = 'NON-MED' and MED_TYPE = '" + bType + "' and PO.FACE_AMOUNT > 100000 and PO.FACE_AMOUNT <= 150000))" +
            // "GROUP BY UR.REGION_CODE,REGION_NAME"+
            //" ORDER BY UR.REGION_CODE";

            //            commPolicy.CommandText =  "select(UR.REGION_CODE)," +
            // "sum(TERM_COUNT) as TERM_COUNT," +
            // "SUM(NET_COUNT) as NET_COUNT, " +
            //                "  sum(PAID_COUNT) as PAID_COUNT" +
            // " from NETREGION UR inner join POLICIES2 PO ON  PO.AGENT_NUMBER = UR.AGENT_NUMBER left outer join U_REQ PB ON PB.AGENT_NUMBER = UR.AGENT_NUMBER" +
            //" WHERE UR.AGENT_NUMBER in " +
            //" (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = 'WRE')  AND " +
            // "PB.COMPANY_CODE = '07' AND UR.REGION_CODE IS NOT NULL AND" +
            // "(UR.PAY_DATE BETWEEN '" + frmDate + "' AND '" + tDate + "') AND" +
            // "((LEFT(UR.REGION_CODE, 1) = '" + RegionCodeAll + "' ) OR(SUBSTRING(UR.REGION_CODE, 2, 1) ='" + RegionCodeAll + "' ) OR(LEFT(UR.REGION_CODE, 2) = '" + RegionCodeAll + "') OR('ALL' = '" + RegionCodeAll + "')) AND((MED_TYPE = '" + bType + "') OR"+
            //"('" + bType + "' = 'ALL') OR ( MED_TYPE = 'NON_MED'and MED_TYPE = '" + bType + "' and PO.FACE_AMOUNT > 100000 and PO.FACE_AMOUNT <= 150000))" +
            // "GROUP BY UR.REGION_CODE,REGION_NAME" +
            //" ORDER BY UR.REGION_CODE";
           // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);


            commPolicy.Parameters.AddWithValue("@agentid", SqlDbType.VarChar).Value = HierAgent;
            commPolicy.Parameters.AddWithValue("@company", SqlDbType.VarChar).Value = sCompany;
            commPolicy.Parameters.AddWithValue("@fromdate", SqlDbType.VarChar).Value = frmDate;
            commPolicy.Parameters.AddWithValue("@todate", SqlDbType.VarChar).Value = tDate;
            commPolicy.Parameters.AddWithValue("@orderby", SqlDbType.VarChar).Value = sOrderBy;
            commPolicy.Parameters.AddWithValue("@orderdir", SqlDbType.VarChar).Value = sOrderDir;
            //commPolicy.Parameters.AddWithValue("@resulttype", SqlDbType.VarChar).Value = sResultType;
            commPolicy.Parameters.AddWithValue("@resulttype", SqlDbType.VarChar).Value = bType;
            commPolicy.Parameters.AddWithValue("@green", SqlDbType.VarChar).Value = IsGreen;



           // commPolicy.CommandTimeout = 0;
            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy);

          //  dataadapter.SelectCommand.CommandTimeout = 3600;
            datatab = new DataTable(); // Create a new Data table
            dataadapter.Fill(datatab);


            //adPolicy.Fill(dsPolicy);
            //if (dsPolicy != null && dsPolicy.Tables[0] != null)
            //    datatab = dsPolicy.Tables[0];

            //commPolicy.ExecuteNonQuery();
            //Fill the data table for export excel
            if (datatab != null)
                dataPolicy = datatab;
            con.Close();

            if (datatab != null && datatab.Rows.Count == 0)
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
                GridView1.DataSource = datatab;
                GridView1.DataBind();
                if (dataPolicy != null && dataPolicy.DefaultView != null)
                    dataPolicy = dataPolicy.DefaultView.ToTable();

                if (datatab != null)
                {
                    lblcount.Text = datatab.Rows.Count.ToString();
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
