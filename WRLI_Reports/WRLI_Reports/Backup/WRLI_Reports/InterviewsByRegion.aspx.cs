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

//using ClosedXML.Excel;

//using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace WRLI_Reports
{
    public partial class InterviewsByRegion : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //string agent = "WRE";
        DataSet dsPolicy = new DataSet();
        string sCompany = "07";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = "UR.REGION_CODE";
        string OrderDir = "ASC";
        bool bNet = false;
        bool bRegion = false;
        int nRowct = 0;
        bool IsLatestRef = false;
         //string sType = "type";
         Int32[] arr_NB = new Int32[] { };
         string[] strRowValue = new string[3];
         DataTable datatab = new DataTable(); // Create a new Data table
         DataTable datatabTotal = new DataTable();
         public static DataTable dataPolicy = new DataTable();
         string RegionCodeAll = "ALL";
         string sType = "'146','176','208','209','210','150','151','152','195','367','368','369','128','97','137','138','135','98','139'";
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
                txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                else
                    sCompany = "07";
                if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                    AgentID = Session["LoginID"].ToString();

                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();
                //If Distributor is set in Session the read the value.
                if ((Session["Distributor"] != null) && ( Session["Distributor"].ToString() == "HMI" || Session["Distributor"].ToString() == "Texas" || Session["Distributor"].ToString() == "NEAT"
                     || Session["Distributor"].ToString() == "MGA"))
                    sType = Session["Distributor"].ToString();
                else
                    sType = "'146','176','208','209','210','150','151','152','195','367','368','369','128','97','137','138','135','98','139'";
                //Read the type value from query string
                if (Request.QueryString["type"] != null)
                {
                    if (Request.QueryString["type"] == "1")
                    {
                        sType = "'195'";
                    }
                    else if (Request.QueryString["type"]== "2" )
                    {
                        sType = "'176'";
                    }
                    else if (Request.QueryString["type"] == "3" )
                    {
                        sType = "'146'";
                    }
                    else if (Request.QueryString["type"] == "4" )
                    {
                        sType = "'208','209','210','150','151','152'";
                    }
                }

                //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Initial Catalog=WRE_AGENT");
               
            }
            
        }

        protected void InvokeSP()
        {

            string[] fromDate = txtFrom.Text.Split('/');
            FromDate = fromDate[2] + fromDate[0] + fromDate[1];

            string[] toDate = txtTo.Text.Split('/');
            ToDate = toDate[2] + toDate[0] + toDate[1];

            con.Open();
            string bType = "ALL";
            if(rdListType.Items[0].Selected)
            {
                bType = "ALL";
            }
            SqlCommand commPolicy = new SqlCommand();

            commPolicy.Connection = con;
            //if (Request.QueryString["LatestRef"] != null && Request.QueryString["LatestRef"] != "")
            //Interview for agent and the else part is to get the latest ref ID 
            // If the check box is selected
            //if (chkInterview.Checked)
            if ((chkInterview.Checked) || (rdListType.SelectedItem != null && rdListType.SelectedItem.Text == "All"))
            {
                commPolicy.CommandText = "SELECT PA.POLICY_NUMBER, pa.agent_number, pa.agentfirst + ' ' +pa.agentlast as agentname, pa.referenceid,pa.insuredfirst + ' ' + pa.insuredlast as "+
" INSURED_NAME, PI.FACE_AMOUNT, PI.CONTRACT_CODE, PI.CONTRACT_DESC, PI.PAID_TO_DATE, PI.ISSUE_DATE, PD.ID, CASE WHEN PD.DIRECTION = 1 THEN 'Outbound'"+
 " ELSE 'Inbound' END AS DIRECTION, PD.STARTTIME, PD.STOPTIME, DATEDIFF(MINUTE,PD.STARTTIME,PD.STOPTIME) as DURATION, PD.HOLDTIME, "+
  " CAST(PA.PRODUCTID AS CHAR(3)) AS PRODID FROM PART_A PA LEFT OUTER JOIN PART_A_DETAIL PD ON PA.ID = PD.ID LEFT OUTER JOIN POLICY PI ON "+
   " PA.COMPANY_CODE = PI.COMPANY_CODE AND PA.POLICY_NUMBER = PI.POLICY_NUMBER WHERE (PA.COMPANY_CODE = '" + sCompany + "' ) AND " +
   " CAST(PA.PRODUCTID AS CHAR(3)) IN (" + sType + ") " +
   " AND CAST(CONVERT(CHAR(10),PA.TIMERECEIVED,101) AS DATETIME) BETWEEN '" + FromDate + "' AND '" + ToDate + "' ORDER BY PA.POLICY_NUMBER ASC";
                //If Latest ref check box is checked.
                if (chkRef.Checked)
                {
                    commPolicy.CommandText = "SELECT PA.POLICY_NUMBER, pa.agent_number, pa.agentfirst + ' ' +pa.agentlast as agentname, pa.referenceid,pa.insuredfirst + ' ' + pa.insuredlast as " +
" INSURED_NAME, PI.FACE_AMOUNT, PI.CONTRACT_CODE, PI.CONTRACT_DESC, PI.PAID_TO_DATE, PI.ISSUE_DATE, PD.ID, CASE WHEN PD.DIRECTION = 1 THEN 'Outbound'" +
 " ELSE 'Inbound' END AS DIRECTION, PD.STARTTIME, PD.STOPTIME, DATEDIFF(MINUTE,PD.STARTTIME,PD.STOPTIME) as DURATION, PD.HOLDTIME, " +
  " CAST(PA.PRODUCTID AS CHAR(3)) AS PRODID FROM PART_A PA LEFT OUTER JOIN PART_A_DETAIL PD ON PA.ID = PD.ID LEFT OUTER JOIN POLICY PI ON " +
   " PA.COMPANY_CODE = PI.COMPANY_CODE AND PA.POLICY_NUMBER = PI.POLICY_NUMBER WHERE (PA.COMPANY_CODE = '" + sCompany + "' ) AND " +
   " CAST(PA.PRODUCTID AS CHAR(3)) IN (" + sType + ") " +
   " AND CAST(CONVERT(CHAR(10),PA.TIMERECEIVED,101) AS DATETIME) BETWEEN '" + FromDate + "' AND '" + ToDate + "' and starttime in " +
" (select top 1 starttime from part_a_detail pb where pd.id = pb.id order by starttime desc) ORDER BY PA.POLICY_NUMBER ASC";
                }
            }
            else
            {
                // If the check box is not selected
                commPolicy.CommandText = " SELECT DISTINCT(HT.REGION_CODE) AS REGION_CODE, ISNULL (CASE WHEN (MC.AGENCY_NAME IS NOT NULL)and(MC.AGENCY_NAME <>'') THEN MC.AGENCY_NAME ELSE MC.AGENT_NAME END,' ') as REGION_NAME, COUNT(DISTINCT PA.ID) as CALL_COUNT, AVG(PD.HOLDTIME) as HOLD_TIME, AVG(DATEDIFF(MINUTE,STARTTIME,STOPTIME)) as DURATION FROM AGENT_HIERLIST HL INNER JOIN AGENTS HT ON HT.AGENT_NUMBER = HL.AGENT_NUMBER and HT.COMPANY_CODE = HL.COMPANY_CODE INNER JOIN PART_A PA ON HT.AGENT_NUMBER = PA.AGENT_NUMBER LEFT OUTER JOIN PART_A_DETAIL PD ON PA.ID = PD.ID INNER JOIN REGION_NAMES as MC ON HT.REGION_CODE = MC.MARKETING_COMPANY LEFT OUTER JOIN POLICY as PO ON PA.COMPANY_CODE = PO.COMPANY_CODE AND PA.POLICY_NUMBER = PO.POLICY_NUMBER WHERE (ISNULL(PA.COMPANY_CODE,'07')= '" + sCompany + "' ) AND HL.HIERARCHY_AGENT = '" + AgentID + "' AND CAST(CONVERT(CHAR(10),PA.TIMERECEIVED,101) AS DATETIME) BETWEEN '" + FromDate + "' AND '" + ToDate + "' AND CAST(PA.PRODUCTID AS CHAR(3)) IN (" + sType + ") GROUP BY HT.REGION_CODE,MC.AGENCY_NAME, MC.AGENT_NAME ORDER BY REGION_CODE ASC";
            }

            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);
            
           // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // Create a SQL Data Adapter and assign it the cmd value. 
            datatab = new DataTable(); // Create a new Data table
            dataadapter.Fill(datatab);
            //Store the results in data table
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
                int nTotCt =  datatab.Rows.Count - 1;
                dvgrid.Visible = true;
                lblcount.Visible = true;
                //Added newly to implement Export to Excel functionality
                if(dataPolicy != null && dataPolicy.DefaultView != null )
                    dataPolicy = dataPolicy.DefaultView.ToTable();

                if ((chkInterview.Checked) || (rdListType.SelectedItem !=null && rdListType.SelectedItem.Text == "All"))
                {
                    grInterviewsByRegion.DataSource = datatab;
                    grInterviewsByRegion.DataBind();
                    
                }
                else
                {
                    AddEditRows();
                }
                lblcount.Text = "Total Record Count: " + nTotCt;
            }

            //grdHandling.DataSource = datatab;
            //grdHandling.DataBind();
            // adPolicy.Fill(dsPolicy);
            commPolicy.Connection = con;
            commPolicy.CommandText = "SELECT COUNT(DISTINCT PA.ID) as CALL_COUNT, AVG(PD.HOLDTIME) as HOLD_TIME, AVG(DATEDIFF(MINUTE,STARTTIME,STOPTIME)) as DURATION FROM " +
 " AGENT_HIERLIST HL INNER JOIN AGENTS HT ON HT.AGENT_NUMBER = HL.AGENT_NUMBER and HT.COMPANY_CODE = HL.COMPANY_CODE INNER JOIN PART_A PA ON " +
 " HT.AGENT_NUMBER = PA.AGENT_NUMBER LEFT OUTER JOIN PART_A_DETAIL PD ON PA.ID = PD.ID LEFT OUTER JOIN POLICY as PO ON " +
 " PA.COMPANY_CODE = PO.COMPANY_CODE AND PA.POLICY_NUMBER = PO.POLICY_NUMBER WHERE (ISNULL(PA.COMPANY_CODE,'07')= '" + sCompany + "' ) AND" +
  " HL.HIERARCHY_AGENT = '" + AgentID + "' AND CAST(CONVERT(CHAR(10),PA.TIMERECEIVED,101) AS DATETIME) BETWEEN '" + FromDate + "' AND '" + ToDate + "' AND " +
  " CAST(PA.PRODUCTID AS CHAR(3)) IN (" + sType + ")";
  //('146','176','208','209','210','150','151','152','195','367','368','369','128','97','137','138','135','98','139') ";
            commPolicy.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter dataadapt = new SqlDataAdapter(commPolicy.CommandText, con);
            // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // Create a SQL Data Adapter and assign it the cmd value. 
            datatabTotal = new DataTable(); // Create a new Data table
            dataadapt.Fill(datatabTotal);
            con.Close();
            
            //commPolicy.ExecuteNonQuery();
            //con.Close();


            
        }

        protected void InitPaidReportColumns(int Rowcount)
        {
            arr_NB[0] = new Int32();
            arr_NB[1] = new Int32();
            arr_NB[2] = new Int32();
            arr_NB[3] = new Int32();

        }

        protected void PRReadColumnValues(DataRowView DataRowCurrView, ref Int32[] arr_NB)
        {
            arr_NB[1] += Convert.ToInt32(DataRowCurrView["Call_Count"]);
            arr_NB[2] += Convert.ToInt32(DataRowCurrView["Hold_Time"]);
            arr_NB[3] += Convert.ToInt32(DataRowCurrView["Duration"]);

        }
        protected void Ytd_Click(object sender, EventArgs e)
        {
            //Response.Redirect("CurrentDateReport_Averages.aspx?GroupYTD=true");
        }

        protected void Mtd_Click(object sender, EventArgs e)
        {
            //Response.Redirect("CurrentDateReport_Averages.aspx?GroupYTD=false");
        }


        protected void Button1_Click(object sender, EventArgs e)
        {
            tblgrid.Visible = true;
            string selectedComp = "ALL";
            string selectedAgent = "ALL";
            InvokeSP();
            

        }

        protected void InitGridColumns(int Rowcount)
        {
            arr_NB[0] = new Int32();
            arr_NB[1] = new Int32();
            arr_NB[2] = new Int32();
            arr_NB[3] = new Int32();
            //arr_NB[4] = new Int32();

        }
        protected void ReadColumnValues(DataRowView DataRowCurrView, ref Int32[] arr_NB)
        {
            arr_NB[1] += Convert.ToInt32(DataRowCurrView["Call_Count"]);
            arr_NB[2] += Convert.ToInt32(DataRowCurrView["Hold_Time"]);
            arr_NB[3] += Convert.ToInt32(DataRowCurrView["Duration"]);

            //arr_NB[0] += Convert.ToInt32(DataRowCurrView["DISPLAYID"]);
            //arr_NB[1] += Convert.ToInt32(DataRowCurrView["DISPLAYNAME"]);
            /*if (DataRowCurrView["Call_Count"] != null && DataRowCurrView["Call_Count"].ToString().Trim() != "")
                strRowValue[0] += DataRowCurrView["Call_Count"] + "~";
            if (DataRowCurrView["Hold_Time"] != null && DataRowCurrView["Hold_Time"].ToString().Trim() != "")
                strRowValue[1] += DataRowCurrView["Hold_Time"] + "~"; */
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
                DataRowCurrView = MyDataView1[nIndex];
                
                ReadColumnValues(DataRowCurrView, ref arr_NB);
                // End of new logic
            }
            MyDataView1.AllowNew = true;
            DataRowView MyDataRowView = MyDataView1.AddNew();
            
            int position = 0;
            int i = 0;
            MyDataView1.AllowEdit = true;
            MyDataRowView.BeginEdit();
            position = i + 1; //Dont want to insert at the row, but after.
            //if (FilterResultsType == "1")
            if("1" == "1")
            {

                MyDataRowView["Region_Name"] = " Total ";
                MyDataRowView["Call_Count"] = arr_NB[1]; 
                MyDataRowView["Hold_Time"] = arr_NB[2]; 
                MyDataRowView["Duration"] = arr_NB[3]; 
                
            }
            MyDataRowView.EndEdit();
            grInterviewsByRegion.DataSource = MyDataView1;
            //grInterviewsByRegion.DataSource = datatab;
            grInterviewsByRegion.DataBind();

        }

        protected void ddlregion_PreRender(object sender, EventArgs e)
        {
            
        }


        protected void ddlstate_PreRender(object sender, EventArgs e)
        {
            
        }


        protected void ddlpolicydesc_PreRender(object sender, EventArgs e)
        {
            
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            //Export the grid values as Excel report
            ExportToExcel();
            //Export("InterviewByRegion.xls", this.grInterviewsByRegion);
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
        
        protected void ExportToExcel_OLD()
        {
                InvokeSP();
             
                if (datatab.Rows.Count > 0 && datatab != null)
            {                
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"Interviews_By_Region" + DateTime.Now.ToString();                
                Response.AddHeader("Content-Disposition", "inline;filename=" + filename.Replace("/", "").Replace(":", "")+".xlsx");
                  //Call  Export function
                //Response.BinaryWrite(ExportToCSVFileOpenXML(datatab));                                
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
