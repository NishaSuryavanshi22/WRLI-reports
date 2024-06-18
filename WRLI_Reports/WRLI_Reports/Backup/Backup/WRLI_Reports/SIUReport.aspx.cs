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

//using ClosedXML.Excel;

//using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace WRLI_Reports
{
    public partial class SIUReport : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
//        string agent = "WRE";
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
         string sType = "type";
         Int32[] arr_NB = new Int32[] { };
         string[] strRowValue = new string[3];
         DataTable datatab = new DataTable(); // Create a new Data table
         string RegionCodeAll = "ALL";
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                tblgrid.Visible = false;
                txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                
                if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                    AgentID = Session["LoginID"].ToString();

                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();
                //

                SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + AgentID + "' ORDER BY SORTNAME ASC");
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
                ddlHandagent.DataBind();

               
            }
        }

        protected void InvokeSP()
        {

            string selectedAgent = "ALL";
            string[] fromDate = txtFrom.Text.Split('/');
            FromDate = fromDate[2] + fromDate[0] + fromDate[1];

            string[] toDate = txtTo.Text.Split('/');
            ToDate = toDate[2] + toDate[0] + toDate[1];

            con.Open();
            string bType = "ALL";
            int indexagent = ddlHandagent.SelectedItem.Value.LastIndexOf("-");
            if (indexagent > 0)
            {
                selectedAgent = ddlHandagent.SelectedItem.Value.Substring(0, indexagent);
            }

            SqlCommand commPolicy = new SqlCommand();

            commPolicy.Connection = con;
            //commPolicy.CommandType = CommandType.StoredProcedure;
            commPolicy.CommandText = " select DISTINCT PA.REFERENCEID, PA.POLICY_NUMBER, dbo.LPDATE_TO_STRDATE(ISNULL(PO.APP_RECEIVED_DATE,'')) AS APP_RECEIVED_DATE, PA.COMPANY_CODE, CASE WHEN (AG.NAME_FORMAT_CODE = 'C') THEN AG.NAME_BUSINESS ELSE RTRIM(AG.INDIVIDUAL_LAST) + ', ' + RTRIM(AG.INDIVIDUAL_FIRST) END AS AGENT_NAME, PA.AGENT_NUMBER, dbo.FORMAT_PHONE2(AG.TELE_NUM_OFFICE) AS TELE_NUM_OFFICE, dbo.FORMAT_PHONE2(AG.TELE_NUM) AS TELE_NUM, dbo.FORMAT_PHONE2(AG.FAX_NUM) AS FAX_NUM, dbo.FORMAT_PHONE2(AG.TELE_NUM_CELL) AS TELE_NUM_CELL, AG.REGION_CODE, CASE WHEN (PO.PI_FORMAT = 'C') THEN PO.PI_BUSINESS ELSE (case when pa.policy_number is NULL then (upper(pa.insuredlast)+', '+upper(pa.insuredfirst)) else RTRIM(PO.PI_LAST) + ', ' + RTRIM(PO.PI_FIRST) end)END AS PI_NAME, CASE WHEN (OWNER_FORMAT = 'C') THEN OWNER_BUSINESS ELSE RTRIM(OWNER_LAST) + ', ' + RTRIM(OWNER_FIRST) END AS OWNER_NAME, RTRIM(PO1.PAY_LAST) + ', ' + RTRIM(PO1.PAY_FIRST) AS PAY_NAME, PI_ADDRESS1,PI_CITY,PO.PI_STATE,PI_ZIP, OWNER_CITY,OWNER_STATE, ISSUE_STATE, CASE WHEN (BEN_FORMAT = 'C') THEN BEN_BUSINESS ELSE RTRIM(BEN_LAST) + ', ' + RTRIM(BEN_FIRST) END AS BEN_NAME, ISNULL(RELATIONSHIP_OF_BENEFICIARY,'') AS BEN_RELATIONSHIP, PO.CONTRACT_CODE,PO.CONTRACT_REASON,PO.CONTRACT_DESC, PO.PLAN_CODE, PO.BILLING_MODE, ISNULL(PO.MODE_PREMIUM,0) AS MODE_PREMIUM, PO.RATE_CLASS, HIERARCHY_AGENT, CALLERIDNAME, dbo.FORMAT_PHONE(CALLERIDNUMBER) AS CALLER_ID, dbo.FORMAT_PHONE(ISNULL(PI_PHONE,'')) AS PI_PHONE, ISNULL(PO1.EFT_ROUTING_NUMBER,ISNULL(B.ROUTING_NUMBER,'')) as EFT_ROUTING_NUMBER, ISNULL(PO1.EFT_ACCOUNT_NUMBER,ISNULL(B.ACCOUNT_NUMBER,'')) AS EFT_ACCOUNT_NUMBER, TIMERECEIVED from PART_A PA left outer join PART_A_DETAIL PD on PA.ID = PD.ID left outer join POLICIES2 PO on PA.POLICY_NUMBER = PO.POLICY_NUMBER AND PA.COMPANY_CODE = PO.COMPANY_CODE left outer join POLICY PO1 on PA.POLICY_NUMBER = PO1.POLICY_NUMBER AND PA.COMPANY_CODE = PO1.COMPANY_CODE LEFT OUTER join BANKING B on PA.POLICY_NUMBER = B.POLICY_NUMBER AND PA.COMPANY_CODE = PO1.COMPANY_CODE LEFT OUTER join PENDING_POLICY PP on PA.POLICY_NUMBER = PP.POLICY_NUMBER AND PA.COMPANY_CODE = PP.COMPANY_CODE LEFT OUTER join AGENTS AG on PA.AGENT_NUMBER = AG.AGENT_NUMBER AND PA.COMPANY_CODE = AG.COMPANY_CODE WHERE ((PO.AGENT_NUMBER = '" + selectedAgent + "' OR 'ALL'='ALL')) AND PA.AGENT_NUMBER IN (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT =  '" + AgentID + "' )AND (TIMERECEIVED BETWEEN dbo.LPDATE_TO_DATE('" + FromDate + "') AND dbo.LPDATE_TO_DATE( '" + ToDate + "' )) and pa.company_code= '" + sCompany + "' ORDER BY POLICY_NUMBER ";
 

            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);
            
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
                AddEditRows();
                lblcount.Text = "Total Policy Count: " + datatab.Rows.Count.ToString();
            }

            //grdHandling.DataSource = datatab;
            //grdHandling.DataBind();

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
            //arr_NB[3] = new Int32();
            //arr_NB[4] = new Int32();

        }
        protected void ReadColumnValues(DataRowView DataRowCurrView, ref Int32[] arr_NB)
        {

            //arr_NB[0] += Convert.ToInt32(DataRowCurrView["DISPLAYID"]);
            //arr_NB[1] += Convert.ToInt32(DataRowCurrView["DISPLAYNAME"]);
            /*if (DataRowCurrView["REGION_NAME"] != null && DataRowCurrView["REGION_NAME"].ToString().Trim() != "")
                strRowValue[0] +=DataRowCurrView["REGION_NAME"] + "~";
            if (DataRowCurrView["REGION_CODE"] != null && DataRowCurrView["REGION_CODE"].ToString().Trim() != "")
                strRowValue[1] += DataRowCurrView["REGION_CODE"] + "~"; */
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
            /* Below line Commented by Siva
            DataRowView MyDataRowView = MyDataView1.AddNew();
            
            //DataRow MyDataRowView = MyDataView1.Table.NewRow();
            //MyDataView1.Table.Rows.InsertAt(MyDataRowView, 0); 

            int position = 0;
            int i = 0;
            MyDataView1.AllowEdit = true;
            MyDataRowView.BeginEdit();
            position = i + 1; //Dont want to insert at the row, but after.
            //if (FilterResultsType == "1")
            if("1" == "1")
            {
                MyDataRowView["DISPLAYID"] = "Total";
                //MyDataRowView["SUB_COUNT"] = arr_NB[3];
                //MyDataRowView["SUB_PREM"] = arr_NB[4];
                MyDataRowView["SUB_PREM"] = "1";
                MyDataRowView["SUB_COUNT"] = "2";
                
            } */
            //grDailyAverages.DataSource = MyDataView1;
            grRequirementsByRegion.DataSource = datatab;
            grRequirementsByRegion.DataBind();

        }


        protected void ddlHandagent_PreRender(object sender, EventArgs e)
        {
            if (ddlHandagent.Items.FindByValue(string.Empty) == null)
            {
                ddlHandagent.Items.Insert(0, new ListItem("ALL", "ALL"));
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
            //ExportToExcel();
            Export("RequirementByRegion.xls", this.grRequirementsByRegion);
        }

        
        
        protected void ExportToExcel()
        {
                InvokeSP();
             
                if (datatab.Rows.Count > 0 && datatab != null)
            {                
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"RequirementByRegion" + DateTime.Now.ToString();                
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
