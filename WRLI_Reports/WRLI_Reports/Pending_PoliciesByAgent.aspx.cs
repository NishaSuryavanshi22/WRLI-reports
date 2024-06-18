﻿using System;
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
using System.Configuration;

namespace WRLI_Reports
{
    public partial class Pending_PoliciesByAgent : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        // SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //string agent = "WRE";
        DataSet dsPolicy = new DataSet();
        SqlCommand commPolicy = new SqlCommand();
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
         public static DataTable dataPolicy = new DataTable();
         string RegionCodeAll = "ALL";
         public string ContractCode = "";
         public string ContractCodeQuery = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                InitiateControls();
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
                tblgrid.Visible = true ;
                txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                else
                
                if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                    AgentID = Session["LoginID"].ToString();

                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();

                SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + AgentID + "' ORDER BY SORTNAME ASC");
                comm.Connection = con;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(comm.CommandText, con);
                ad.Fill(ds);

                List<string> lstagent = new List<string>();
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    lstagent.Add(ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() + "-" + ds.Tables[0].Rows[i].ItemArray[1].ToString().Trim());

                }

                ddlAgentList.DataSource = lstagent;
                ddlAgentList.DataBind();
               
            }
            //InitiateControls();
        }

        protected void InvokeSP()
        {

            string[] fromDate;
            string[] toDate;
            if (txtFrom.Text.Contains('/'))
            {
                fromDate = txtFrom.Text.Split('/');
                FromDate = fromDate[2] + fromDate[0] + fromDate[1];
                //
            }
            else if (txtFrom.Text.Contains('-'))
            {
                fromDate = txtFrom.Text.Split('-');
                FromDate = fromDate[2] + fromDate[0] + fromDate[1];
            }
            if (txtTo.Text.Contains('/'))
            {
                toDate = txtTo.Text.Split('/');
                ToDate = toDate[2] + toDate[0] + toDate[1];
            }

            else if (txtTo.Text.Contains('-'))
            {
                toDate = txtTo.Text.Split('-');
                ToDate = toDate[2] + toDate[0] + toDate[1];
            }

            //if (ContractCode == "ALL")
            //{
            //    // ContractCodeQuery = "((PO.CONTRACT_CODE = 'A')OR(PO.CONTRACT_CODE = 'T')OR(PO.CONTRACT_CODE = 'P')OR(PO.CONTRACT_CODE = 'S')) ";
            //    //NISHA
            //    ContractCodeQuery = "((PR.CONTRACT_CODE = 'A')OR(PR.CONTRACT_CODE = 'T')OR(PR.CONTRACT_CODE = 'P')OR(PR.CONTRACT_CODE = 'S')) ";

            //}

            //string[] fromDate = txtFrom.Text.Split('/');
            //FromDate = fromDate[2] + fromDate[0] + fromDate[1];

            //string[] toDate = txtTo.Text.Split('/');
            //ToDate = toDate[2] + toDate[0] + toDate[1];

            con.Open();
            string bType = "ALL";
            string sAgentList = "";
             commPolicy = new SqlCommand();
            sAgentList = ddlAgentList.SelectedItem.Value;
            string[] SelAgentID = sAgentList.Split('-');

            if (SelAgentID != null && SelAgentID.Length > 0)
            {
                AgentID = SelAgentID[0].ToString();
            }
            
            commPolicy.Connection = con;

            commPolicy.CommandText = "SELECT DISTINCT PR.POLICY_NUMBER, PR.CONTRACT_CODE, B.PRODUCT_CODE AS PRODUCT_ID, I.PLAN_CODE,H.LINE_OF_BUSINESS,H.ISSUE_STATE,  B.CWA_TRANSFER_AMT AS CASH_WITH_APPL,B.FACE_AMOUNT, PR.AGENT_NUMBER as AGENT_NUMBER, '' as MARKET, '' AS LEVEL, PR.RECORD_TYPE as REC_TYPE, PR.REGION_CODE as REGION, st.prod_pcnt as policy_split,H.MODE_PREMIUM, " +
"  (RTRIM(ISNULL(PR.PI_LAST,''))+','+RTRIM(ISNULL(PR.PI_FIRST,''))+RTRIM(ISNULL(H.PI_MIDDLE,''))) AS INSURED_FULL_NAME,I.ISSUE_AGE, H.SA_REGION_CODE AS AGENCY_CODE, H.SA_REGION_CODE AS SERVICING_AGENCY," +
 "H.BILLING_MODE,(H.MODE_PREMIUM * 12 / H.BILLING_MODE) AS ANNZD_PREMIUM, "+
  "B.PAYMENT_FLAG, SUBSTRING(CAST(B.PAYMENT_DATE AS CHAR),5,2)+'/'+SUBSTRING(CAST(B.PAYMENT_DATE AS CHAR),7,2)+'/'+SUBSTRING(CAST(B.PAYMENT_DATE AS CHAR),1,4) as "+
  " PAYMENT_DATE,"+
   " CASE WHEN RTRIM(HL.NAME_BUSINESS)<>'' THEN HL.NAME_BUSINESS ELSE RTRIM(HL.INDIVIDUAL_LAST) +', '+RTRIM(HL.INDIVIDUAL_FIRST) END AS AGENT_NAME,"+
    " H.MAX_TARGET,c2.description FROM POLICIES2  PR WITH (NOLOCK) LEFT OUTER JOIN PENDING_POLICY as B "+
     " WITH (NOLOCK) ON (B.COMPANY_CODE=PR.COMPANY_CODE AND B.POLICY_NUMBER=PR.POLICY_NUMBER) LEFT OUTER JOIN "+
     " REGION_NAMES as MC WITH (NOLOCK) ON (PR.REGION_CODE = MC.MARKETING_COMPANY) LEFT OUTER JOIN POLICY AS H "+
     " WITH (NOLOCK) ON (H.COMPANY_CODE = PR.COMPANY_CODE AND H.POLICY_NUMBER=PR.POLICY_NUMBER) LEFT OUTER JOIN POLICY_COVERAGE AS I "+
     " WITH (NOLOCK) ON (I.COMPANY_CODE = PR.COMPANY_CODE AND I.POLICY_NUMBER = PR.POLICY_NUMBER AND I.PLAN_CODE IS NOT NULL) "+
     " LEFT OUTER JOIN POLICY_SPLIT ST WITH (NOLOCK) ON (PR.COMPANY_CODE = ST.COMPANY_CODE) and (PR.POLICY_NUMBER = ST.POLICY_NUMBER) AND "+
      " (PR.AGENT_NUMBER = ST.AGENT_NUMBER) and (PR.ISSUE_DATE = ST.ISSUE_DATE) left outer JOIN AGENT_HIERLIST HL WITH (NOLOCK) "+
      " ON PR.COMPANY_CODE = HL.COMPANY_CODE AND  PR.AGENT_NUMBER = HL.AGENT_NUMBER AND HL.AGENT_NUMBER = HL.HIERARCHY_AGENT "+
      " LEFT OUTER JOIN PERSISTANCY2 PS WITH (NOLOCK) ON PR.AGENT_NUMBER = PS.AGT_NUM and ((dbo.GETOLDREGION (PR.REGION_CODE) = PS.REGION) "+
      " or (PR.REGION_CODE = PS.REGION)) left outer join coverage2 c2 on i.coverage_id = c2.coverage_id and i.company_code = c2.company_code "+
      " WHERE PR.COMPANY_CODE='15' AND ((PR.PAYMENT_DATE BETWEEN '" + FromDate + "' AND '" + ToDate + "') OR " +
      " (PR.LAST_CHANGE_DATE BETWEEN  '" + FromDate + "' AND '" + ToDate + "') OR (PR.APP_RECEIVED_DATE " +
      " BETWEEN '" + FromDate + "' AND '" + ToDate + "')) and (B.app_received_date BETWEEN  '" + FromDate + "' AND '" + ToDate + "' or " +
       " (b.app_Received_date  is null and b.createddatetime between  '" + FromDate + "' AND '" + ToDate + "' ))  AND "+ ContractCodeQuery + " AND PR.AGENT_NUMBER IN " +
       " (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE COMPANY_CODE= '" + sCompany + "' AND HIERARCHY_AGENT = '" + AgentID + "') " +
         " ORDER BY REGION";


            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);
            dataadapter.SelectCommand = commPolicy;
            //dataadapter.SelectCommand = commPolicy;
           // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // Create a SQL Data Adapter and assign it the cmd value. 
            datatab = new DataTable(); // Create a new Data table
            dataadapter.Fill(datatab);
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

                dvgrid.Visible = true;
                lblcount.Visible = true;
                //AddEditRows();
                lblcount.Text = "Total Policy Count: " + datatab.Rows.Count.ToString();
            }

            grdSubmittedReport.DataSource = datatab;
            grdSubmittedReport.DataBind();

            // adPolicy.Fill(dsPolicy);

            con.Close();
        }

        protected void Ytd_Click(object sender, EventArgs e)
        {
            //Response.Redirect("CurrentDateReport_Averages.aspx?GroupYTD=true");
        }

        protected void Mtd_Click(object sender, EventArgs e)
        {
            //Response.Redirect("CurrentDateReport_Averages.aspx?GroupYTD=false");
        }


        //protected void Button1_Click(object sender, EventArgs e)
        //{
        //    tblgrid.Visible = true;
        //    string selectedComp = "ALL";
        //    string selectedAgent = "ALL";
        //    InvokeSP();
            

        //}


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
        // New changes for Report 28/02/2017
        protected void Button2_Click(object sender, EventArgs e)
        {
            // New changes for Report 28/02/2017
            ExportToExcel();
            //Export("Pending_PoliciesByAgent.xls", this.grdSubmittedReport);
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


        protected void InitiateControls()
        {
            //this.Button2.Click +=new EventHandler(Button2_Click);
            this.All.Click +=new EventHandler(All_Click);
            this.Active.Click+=new EventHandler(Active_Click);
            this.Pending.Click+=new EventHandler(Pending_Click);
            this.Terminated.Click+=new EventHandler(Terminated_Click);
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

        protected void All_Click(object sender, EventArgs e)
        {
            ContractCode = ((System.Web.UI.WebControls.Button)(sender)).Text;
            GetAgentPolicies();
            InvokeSP();

        }

        protected void Active_Click(object sender, EventArgs e)
        {
            ContractCode=((System.Web.UI.WebControls.Button)(sender)).Text;
            GetAgentPolicies();
            InvokeSP();
        }
        protected void Pending_Click(object sender, EventArgs e)
        {
            ContractCode = ((System.Web.UI.WebControls.Button)(sender)).Text;
            GetAgentPolicies();
            InvokeSP();

        }
        protected void Terminated_Click(object sender, EventArgs e)
        {
            ContractCode = ((System.Web.UI.WebControls.Button)(sender)).Text;
            GetAgentPolicies();
            InvokeSP();

        }

        private void GetAgentPolicies()
        {
            string sContractCode = string.Empty;
            if (ContractCode == "ALL")
            {
                // ContractCodeQuery = "((PO.CONTRACT_CODE = 'A')OR(PO.CONTRACT_CODE = 'T')OR(PO.CONTRACT_CODE = 'P')OR(PO.CONTRACT_CODE = 'S')) ";
                //NISHA
                ContractCodeQuery = "((PR.CONTRACT_CODE = 'A')OR(PR.CONTRACT_CODE = 'T')OR(PR.CONTRACT_CODE = 'P')OR(PR.CONTRACT_CODE = 'S')) ";

            }
            else if (ContractCode == "Active")
            {
                // ContractCodeQuery = "(PO.CONTRACT_CODE = 'A')";
                //NISHA

                ContractCodeQuery = "(PR.CONTRACT_CODE = 'A')";

            }
            else if (ContractCode == "Pending")
            {
                //ContractCodeQuery = "(PO.CONTRACT_CODE = 'P')";
                //NISHA

                ContractCodeQuery = "(PR.CONTRACT_CODE = 'P')";
            }
            else if (ContractCode == "Terminated")
                //ContractCodeQuery = "(PO.CONTRACT_CODE = 'T')";
                //NISHA

                ContractCodeQuery = "(PR.CONTRACT_CODE = 'T')";

        }

        private void FetchAndBindValues()
        {
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
                int nTotCt = datatab.Rows.Count - 1;
                dvgrid.Visible = true;
                lblcount.Visible = true;
                //Added newly to implement Export to Excel functionality
                if (dataPolicy != null && dataPolicy.DefaultView != null)
                    dataPolicy = dataPolicy.DefaultView.ToTable();

                //if ((chkInterview.Checked) || (rdListType.SelectedItem !=null && rdListType.SelectedItem.Text == "All"))
                //{
                grdSubmittedReport.DataSource = datatab;
                grdSubmittedReport.DataBind();

                //}
                //else
                //{
                //    AddEditRows();
                //}
                lblcount.Text = "Total Record Count: " + nTotCt;
            }

        }
    }

    
       

}