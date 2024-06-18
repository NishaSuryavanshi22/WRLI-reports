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
using DocumentFormat.OpenXml.Packaging;

//using ClosedXML.Excel;

//using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace WRLI_Reports
{
    public partial class PendingRequirements : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        // SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //string agent = "WRE";
        SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 1000");

        DataSet dsPolicy = new DataSet();
        
        DataTable dataPolicy = new DataTable();
        //string agent = "WRE";
        Int32[] arr_NB = new Int32[] { };
        string sCompany = "15";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = "REGION_CODE";
        string OrderDir = "ASC";
        bool bNet = false;
        bool bRegion = false;
        int nRowct = 0;
         string sType = "type";
         //Int32[] arr_NB = new Int32[] { };
         string[] strRowValue = new string[3];
         DataTable datatab = new DataTable(); // Create a new Data table
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

            if (! IsPostBack)
            {
                tblgrid.Visible = false;
                txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                else
                
                if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                    AgentID = Session["LoginID"].ToString();

                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();

                /* SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + AgentID + "' ORDER BY SORTNAME ASC");
                comm.Connection = con;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(comm.CommandText, con);
                ad.Fill(ds);

                List<string> lstagent = new List<string>();
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    lstagent.Add(ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() + "-" + ds.Tables[0].Rows[i].ItemArray[1].ToString().Trim());

                } */

                //ddlAgentList.DataSource = lstagent;
                //ddlAgentList.DataBind();

                SqlCommand commPolicy = new SqlCommand();
                //15JAN2019 - Siva
                GetComboAgentValues(commPolicy);
            }
        }

        protected void InvokeSP()
        {

            string[] fromDate = txtFrom.Text.Split('/');
            FromDate = fromDate[2] + fromDate[0] + fromDate[1];

            string[] toDate = txtTo.Text.Split('/');
            ToDate = toDate[2] + toDate[0] + toDate[1];
            AgentID = "ALL";
            con.Open();
            string bType = "ALL";
            string sAgentList = "";
            SqlCommand commPolicy = new SqlCommand();
            sAgentList = ddlAgentList.SelectedItem.Value;
            string[] SelAgentID = sAgentList.Split('-');

            if (SelAgentID != null && SelAgentID.Length == 2 )
            {
                AgentID = SelAgentID[1].ToString();
            }
            
            commPolicy.Connection = con;
            string sSort = " ASC";
            /* commPolicy.CommandText = "SELECT DISTINCT H.POLICY_NUMBER, CASE WHEN H.PI_FORMAT = 'C' THEN H.PI_BUSINESS ELSE RTRIM(H.PI_FIRST) + ' ' + RTRIM(H.PI_LAST) END AS IN_FULL_NAME,  H.PRODUCT_CODE AS PRODUCT_ID, PO.LINE_OF_BUSINESS, PO.ISSUE_STATE, H.FACE_AMOUNT, (H.MODE_PREMIUM * 12 / H.BILLING_MODE) AS ANNZD_PREMIUM, H.MODE_PREMIUM, H.BILLING_MODE,  H.COMPANY_CODE,  H.AGENT_NUMBER AS AGENT_NUMBER,CASE WHEN RTRIM(ISNULL(A.NAME_BUSINESS,''))='' THEN RTRIM(A.INDIVIDUAL_LAST)+', '+RTRIM(A.INDIVIDUAL_FIRST) ELSE NAME_BUSINESS END AS AGENT_NAME, A.AGENT_LEVEL AS AGENT_LEVEL, H.REGION_CODE AS REGION,H.RECORD_TYPE as REC_TYPE,H.CONTRACT_CODE, H.PAYMENT_FLAG,H.PAYMENT_DATE"+
            //SUBSTRING(CAST(H.PAYMENT_DATE AS CHAR),5,2)+'/'+ SUBSTRING(CAST(H.PAYMENT_DATE AS CHAR),7,2)+'/'+ SUBSTRING(CAST(H.PAYMENT_DATE AS CHAR),1,4) as H.PAYMENT_DATE "+
" FROM POLICIES2 AS H WITH (NOLOCK) LEFT OUTER JOIN POLICY PO ON PO.POLICY_NUMBER = H.POLICY_NUMBER AND PO.COMPANY_CODE = H.COMPANY_CODE INNER JOIN AGENTS A ON H.COMPANY_CODE = A.COMPANY_CODE AND "+
            " H.AGENT_NUMBER = A.AGENT_NUMBER WHERE (H.COMPANY_CODE='"+  sCompany +  "') AND H.PAYMENT_FLAG='Y' AND H.PAYMENT_DATE >=" + FromDate + " AND H.PAYMENT_DATE <=" + ToDate + " AND H.AGENT_NUMBER IN (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = '"+ AgentID +"')  ORDER BY REGION"; */

            //Changes 15JAN19
            
            /* commPolicy.CommandText = "SELECT DISTINCT PP.*,PP.AGENT_NUMBER as AGENTNUMBER,HL.DISPLAYNAME AS AGENTNAME,PP.INDIVIDUAL_FIRST as FIRSTNAME, PP.INDIVIDUAL_LAST as LASTNAME, PP.POLICY_NUMBER as POLICYNUMBER,RIGHT(RTRIM(PP.SOC_SEC_NUMBER),4) AS SSN,dbo.GET_BASE_PRODUCT_DESC(PP.PRODUCT_CODE) AS PRODUCT_CODE, CASE WHEN PP.DELIVERY_DATE IS NOT NULL THEN 'Policy Mailed' ELSE PP.WFWORKSTEPNAME END AS WFWORKSTEPNAME, PP.FACE_AMOUNT, PP.APP_RECEIVED_DATE FROM PENDING_POLICY PP INNER JOIN AGENT_HIERLIST HL ON HL.AGENT_NUMBER = PP.AGENT_NUMBER AND HL.COMPANY_CODE = PP.COMPANY_CODE WHERE  ((DATEDIFF(day,dbo.LPDATE_TO_DATE(PP.APP_RECEIVED_DATE), GETDATE()) <= 4))" + " AND HL.HIERARCHY_AGENT = '" + AgentID +  "'" ;   */
                //ORDER BY " + "'" + sSort + "'";  
            
            
            //10 Sep 2019 - Development in progress - Siva start
            commPolicy.CommandText = "";

            // Siva- End

            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);
            dataadapter.SelectCommand = commPolicy;
            //dataadapter.SelectCommand = commPolicy;
           // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // Create a SQL Data Adapter and assign it the cmd value. 
            datatab = new DataTable(); // Create a new Data table
            dataadapter.Fill(datatab);
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

                dvgrid.Visible = true;
                lblcount.Visible = true;
                //AddEditRows();
                lblcount.Text = "Total Policy Count: " + datatab.Rows.Count.ToString();
                dataPolicy = datatab;
            }

            grdSubmittedReport.DataSource = datatab;
            grdSubmittedReport.DataBind();

            // adPolicy.Fill(dsPolicy);

            con.Close();
        }

        protected void GetComboAgentValues(SqlCommand comm)
        {
            comm = new SqlCommand("SELECT top 1000 AGENT_NUMBER,ISNULL(DISPLAYNAME,'') AS DISPLAYNAME,SORTNAME FROM AGENT_HIERLIST");
            comm.Connection = con;
            DataSet ds = new DataSet();
            SqlDataAdapter ad = new SqlDataAdapter(comm.CommandText, con);
            ad.Fill(ds);
            if (ds != null)
                ds.Tables[0].DefaultView.Sort = "SORTNAME ASC";

            List<string> lstagent = new List<string>();
            List<string> lstagentName = new List<string>();



            //ddlagentNameList.Items.Add("ALL");
            //18MAR18- Siva
            //ddlagentNameList.Items.Clear();

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string sName = ds.Tables[0].Rows[i].ItemArray[1].ToString().Trim();
                if (sName != "")
                {
                    //lstagent.Add(ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() + "-" + sName + " (" + ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() + ")");
                    //lstagent.Add(sName + " (" + ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() + ")");
                    lstagent.Add(sName + " -" + ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() );


                    //Adding the Agent name to the dropdown
                    //ddlagentNameList.Items.Add(sName + "-" + ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim());

                    //ddlAgentList.Items.Add(sName + "-" + ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim());

                }
            }
            //Bind the agent name and code to the list.
            ddlAgentList.DataSource = lstagent;
            ddlAgentList.DataBind();
            if (ddlAgentList != null && ddlAgentList.Items.FindByValue("ALL") == null)
            {
                ddlAgentList.Items.Insert(0, new ListItem("ALL", "ALL"));


            }
        }

        protected void Ytd_Click(object sender, EventArgs e)
        {
            //Response.Redirect("CurrentDateReport_Averages.aspx?GroupYTD=true");
        }

        protected void Mtd_Click(object sender, EventArgs e)
        {
            //Response.Redirect("CurrentDateReport_Averages.aspx?GroupYTD=false");
        }


        protected void GO_Click(object sender, EventArgs e)
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
            //arr_NB[3] = new Int32();
            //arr_NB[4] = new Int32();

        }
        protected void ReadColumnValues(DataRowView DataRowCurrView, ref Int32[] arr_NB)
        {

            arr_NB[0] += Convert.ToInt32(DataRowCurrView["PAID_COUNT"]);
            arr_NB[1] += Convert.ToInt32(DataRowCurrView["PAID_PREM"]);
            if (DataRowCurrView["PAID_COUNT"] != null && DataRowCurrView["PAID_COUNT"].ToString().Trim() != "")
                strRowValue[0] += DataRowCurrView["PAID_COUNT"] + "~";
            if (DataRowCurrView["PAID_PREM"] != null && DataRowCurrView["PAID_PREM"].ToString().Trim() != "")
                strRowValue[1] += DataRowCurrView["PAID_PREM"] + "~";
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
                //strRowValue[nIndex] = new String();
                DataRowCurrView = MyDataView1[nIndex];
                //List<Int32> stringList = new List<Int32>();
                ReadColumnValues(DataRowCurrView, ref arr_NB);
                // End of new logic
            } 
            // Below line Commented by Siva
            //DataRowView MyDataRowView = MyDataView1.AddNew();
            
            DataRow MyDataRowView = MyDataView1.Table.NewRow();
            //MyDataView1.Table.Rows.InsertAt(MyDataRowView, 0);
            MyDataView1.Table.Rows.InsertAt(MyDataRowView,nRowct);
            MyDataView1.Table.Columns[0].Caption = "Test";

            int position = 0;
            int i = 0;
            MyDataView1.AllowEdit = true;
            MyDataRowView.BeginEdit();
            position = i + 1; //Dont want to insert at the row, but after.
            //if (FilterResultsType == "1")
            if("1" == "1")
            {
                MyDataRowView["DISPLAYNAME"] = "TOTAL";
                
                //MyDataRowView["SUB_COUNT"] = arr_NB[3];
                //MyDataRowView["SUB_PREM"] = arr_NB[4];
                if ( arr_NB != null  && arr_NB.Length != 0 )
                MyDataRowView["SUB_PREM"] = arr_NB[0].ToString();
                MyDataRowView["SUB_COUNT"] = arr_NB[1].ToString();
                
            }
            grdSubmittedReport.DataSource = MyDataView1;
            //DataRowCurrView.EndEdit();
            grdSubmittedReport.DataSource = datatab;
            grdSubmittedReport.DataBind();

        }


        protected void ddlagent_PreRender(object sender, EventArgs e)
        {
            if (ddlAgentList.Items.FindByValue(string.Empty) == null)
            {
                if(! IsPostBack)
                    ddlAgentList.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
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
            ExportToExcel();
            //Export("PendingRequirements.xls", this.grdSubmittedReport);
        }



        protected void ExportToExcel()
        {
            //InvokeSP();
            //            dataPolicy = Session["dataPolicy"] as DataTable;
            //dataPolicy.DefaultView.Sort = "H.POLICY_NUMBER";
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
        
        protected void ExportToExcel_Old()
        {
                InvokeSP();
             
                if (datatab.Rows.Count > 0 && datatab != null)
            {                
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"Reports3P" + DateTime.Now.ToString();                
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
