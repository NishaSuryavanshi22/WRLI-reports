using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CSCUtils;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace WRLI_Reports
{
    public partial class ClientList : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //  SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);


        DataSet dsPolicy = new DataSet();
        DataTable datatab = new DataTable(); // Create a new Data table
        public static DataTable dataPolicy = new DataTable();
        string sCompany = "15";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = " ";
        string OrderDir = "ASC";
        string sState = "ALL";
        string sStatus = "ALL";
      //  String Agent_num = string.Empty;
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
                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();
              //  tblgrid.Visible = false;
                //fill data
                FillData();

            }
        }


        protected void grdHandling_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            grdHandling.PageIndex = e.NewPageIndex;
            FillData();

        }
        private void FillData()
        {
            tblgrid.Visible = true;

            if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                sCompany = Session["CompanyCode"].ToString();
            if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                AgentID = Session["LoginID"].ToString();

            con.Open();
            string SQLcmd = "select CASE WHEN (PI_FORMAT = 'B') THEN PI_BUSINESS ELSE " +
                            "RTRIM(LTRIM(PI_LAST)) + ', ' + RTRIM(LTRIM(PI_FIRST)) END AS DISPLAYNAME, " +
                            "PI_PHONE , POLICY_NUMBER, " +
                            "CASE CONTRACT_CODE " +
                            "WHEN('A') THEN 'Active' " +
                            "WHEN('P') THEN 'Pending' " +
                            "WHEN('S') THEN 'Suspended' " +
                            "WHEN('T') THEN 'Terminated' " +
                            "END AS STATUS, " +
                            "SA_REGION_CODE, " +
                            "SERVICE_AGENT as AGENT_NUMBER " +
                            "from POLICY " +
                            "where (COMPANY_CODE = '" + sCompany + "') " +
                            "AND(SERVICE_AGENT " +
                            "in " +
                            "(SELECT AGENT_NUMBER from AGENT_HIERLIST where COMPANY_CODE ='"+ sCompany + "' " +
                            "AND HIERARCHY_AGENT = '" +AgentID+ "')) " +
                            " ORDER BY DISPLAYNAME";

            SqlDataAdapter da = new SqlDataAdapter(SQLcmd, con);
            DataTable dt = new DataTable();
                da.Fill(dt);
                grdHandling.DataSource = dt;
                grdHandling.DataBind();
            LBLPolicyCount.Text = dt.Rows.Count.ToString();
               // dsPolicy.Tables[0].Rows.Count.ToString();
                con.Close();
          
        }

        protected void grdHandling_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {

                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                
                string Agent_num = e.Row.Cells[5].Text;
                string Policy_num = e.Row.Cells[2].Text;
                if ((e.Row.RowType == DataControlRowType.DataRow)|| (e.Row.RowType== DataControlRowType.Header))
                {
                    e.Row.Cells[e.Row.Cells.Count - 1].Visible = false;
                    grdHandling.HeaderRow.Cells[e.Row.Cells.Count - 1].Visible = false;

                }
                e.Row.Cells[2].ToolTip = "click to view details";

                string text = e.Row.Cells[2].Text;
                HyperLink link = new HyperLink();
                link.NavigateUrl = "PolicyView.aspx?POLICY_NUMBER=" + Policy_num + "&COMPANY_CODE=" + sCompany + "&AGENT_NUMBER=" + Agent_num + "";
                link.Text = text;
                link.Target = "_blank";
                e.Row.Cells[2].Controls.Add(link);

               //PolicyView.aspx?POLICY_NUMBER=W861860006&COMPANY_CODE=16&AGENT_NUMBER=ID01001
               //e.Row.Cells[2].Text = Convert.ToString("<a href=\"PolicyView.aspx?POLICY_NUMBER="+Policy_num+"&COMPANY_CODE="+sCompany+"&AGENT_NUMBER="+Agent_num+"Target="+"_blank"+" \">"+Policy_num+"</a>");
            }
           
        }

        protected void grdHandling_SelectedIndexChanged(object sender, EventArgs e)
        {
            //do nothing
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
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
    }
}