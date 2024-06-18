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
    public partial class ReturnsByRegion : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
//        SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
       // SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //string agent = "WRE";
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
        bool bNet = false;
        bool bRegion = false;
        int nRowct = 0;
        bool IsLatestRef = false;
         string sType = "";
         Int32[] arr_NB = new Int32[] { };
         string[] strRowValue = new string[3];
         //DataTable datatab = new DataTable(); // Create a new Data table
         DataTable datatabTotal = new DataTable();
         string RegionCodeAll = "ALL";
         //string sType = "'146','176','208','209','210','150','151','152','195','367','368','369','128','97','137','138','135','98','139'";
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
                
                
                //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Initial Catalog=WRE_AGENT");
               
            }
            
        }

        protected void InvokeSP()
        {
            string sPaidColumn ="PAYMENT_DATE";
            string StartMonth = "";
            string StartYear = "";
            string EndMonth = "";
            string EndYear = "";
            string[] sFromDate = txtFrom.Text.Split('/');
            string[] sToDate = txtTo.Text.Split('/');
            if (sFromDate != null)
            {
                FromDate = sFromDate[2] + sFromDate[0] + sFromDate[1];
                ToDate = sToDate[2] + sToDate[0] + sToDate[1];
                StartMonth = sFromDate[0];
                StartYear = sFromDate[2];
                EndMonth = sToDate[0];
                EndYear = sToDate[2];


            }

            con.Open();
            string bType = "ALL";
            if(rdListType.Items[0].Selected)
            {
                bType = "ALL";
            }
            SqlCommand commPolicy = new SqlCommand();
            
            commPolicy.Connection = con;
            //Below is the original code 06/Sep/2016 - Should be uncommented before checkin
            commPolicy.CommandText = "SELECT DISTINCT PR.REGION_CODE, COUNT(RC.ENTRY_DATE) as RETURNED_ITEMS, SUM(RC.AMOUNT) as RETURNED_AMOUNT FROM "+ 
                    " RETURN_CHECKS RC INNER JOIN POLICY_HIERARCHY PR ON RC.COMPANY_CODE = PR.COMPANY_CODE AND RC.POLICY_NUMBER = "+
                    " PR.POLICY_NUMBER AND ((RC.SERVICING_AGENT = PR.AGENT_NUMBER)or(RC.WRITING_AGENT = PR.AGENT_NUMBER)) INNER JOIN "+
                    " POLICY_INFO PIN ON PR.COMPANY_CODE = PIN.COMPANY_CODE AND PR.POLICY_NUMBER = PIN.POLICY_NUMBER WHERE "+
                    " ((PR.HIERARCHY_AGENT = 'WRE') ) AND DATEPART(yyyy,ENTRY_DATE) = " + StartYear  + " AND DATEPART(mm,ENTRY_DATE) = "+  StartMonth + " GROUP BY PR.REGION_CODE ORDER BY PR.REGION_CODE ASC "; 

            //Testing code
            /*commPolicy.CommandText = " SELECT DISTINCT PR.Region_Code, PR.AGENT_LEVEL as Returned_Items, PR.Company_Code as Returned_Count FROM " +
                    " POLICY_INFO PI INNER JOIN POLICY_HIERARCHY PR ON PI.COMPANY_CODE = PR.COMPANY_CODE AND PI.POLICY_NUMBER = " +
                    " PR.POLICY_NUMBER WHERE PR.HIERARCHY_AGENT = 'WRE' ";  */ 
            
            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);
            
           // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // Create a SQL Data Adapter and assign it the cmd value. 
            datatab = new DataTable(); // Create a new Data table
            dataadapter.Fill(datatab);
            SqlDataAdapter adPolicy = new SqlDataAdapter(commPolicy.CommandText, con);
            adPolicy.Fill(dsPolicy);
            //
            if (dsPolicy != null && dsPolicy.Tables[0] != null)
                datatab = dsPolicy.Tables[0];
            //Fill the data table for export excel
            if (datatab != null)
                dataPolicy = datatab;
            //
            DataTable dt = new DataTable();
            DataColumn dc = new DataColumn();
            if (datatab != null && datatab.Rows.Count == 0)
            {

                if (dt.Rows.Count == 0)
                {
                    lblcount.Text = "No Records Found for the selected criteria !!!";
                    dvgrid.Visible = false;


                    dt.Columns.Add("REGION_CODE", typeof(string));
                    dt.Columns.Add("RETURNED_ITEMS", typeof(string));
                    dt.Columns.Add("RETURNED_AMOUNT", typeof(string));


                }

                dvgrid.Style.Add("height", "120px");
                DataRow NewRow = dt.NewRow();
                dt.Rows.Add(NewRow);
                grInterviewsByRegion.DataSource = dt;
                grInterviewsByRegion.DataBind();
                lblcount.Text = "No Records Found for the selected criteria !!";
            }
            else
            {
                dvgrid.Style.Add("height", "600px");
                grInterviewsByRegion.DataSource = datatab;
                grInterviewsByRegion.DataBind();
                lblcount.Text = "Total Count = " + dsPolicy.Tables[0].Rows.Count.ToString();
                if (dataPolicy != null && dataPolicy.DefaultView != null)
                    dataPolicy = dataPolicy.DefaultView.ToTable();
            }
            
            con.Close();
            
           
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
            /*arr_NB[1] += Convert.ToInt32(DataRowCurrView["Call_Count"]);
            arr_NB[2] += Convert.ToInt32(DataRowCurrView["Hold_Time"]);
            arr_NB[3] += Convert.ToInt32(DataRowCurrView["Duration"]); */

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
            arr_NB[1] += Convert.ToInt32(DataRowCurrView["Company_Code"]);
            arr_NB[2] += Convert.ToInt32(DataRowCurrView["Sub_Count"]);
            arr_NB[3] += Convert.ToInt32(DataRowCurrView["Paid_Count"]);

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
            DataView MyDataView2 = new DataView(datatabTotal);
            DataRowView DataRowCurrView = null;
            MyDataView1.AllowNew = true;

            //MyDataRowView["active"] = 111;
            //MyDataRowView["sub"] = 222;

            nRowct = datatab.Rows.Count ;
            int nColCt = datatab.Columns.Count;
            arr_NB = new Int32[nColCt];
            if ("1" == "1")
                InitGridColumns(nColCt);
            MyDataView1.AllowNew = true;
            DataRowView MyDataRowView = MyDataView1.AddNew();
            
            int position = 0;
            int i = 0;
            MyDataView1.AllowEdit = true;
            MyDataRowView.BeginEdit();

            for (int nIndex = 0; nIndex < 1 ; nIndex++)
            {
                DataRowCurrView = MyDataView2[nIndex];

                //ReadColumnValues(DataRowCurrView, ref arr_NB);
                // End of new logic

                position = i + 1; //Dont want to insert at the row, but after.
                //if (FilterResultsType == "1")

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
            //ExportToExcel();
            Export("ReturnsByRegion.xls", this.grInterviewsByRegion);
        }

        
        
        protected void ExportToExcel()
        {
                InvokeSP();
             
                if (datatab.Rows.Count > 0 && datatab != null)
            {                
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"Returns_By_Region" + DateTime.Now.ToString();                
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
