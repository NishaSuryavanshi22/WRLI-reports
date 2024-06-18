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
using System.Configuration;

namespace WRLI_Reports
{
    public partial class LapsesByRegion : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection("Data Source=WREMWEBNORR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //string agent = "WRE";
        DataSet dsPolicy = new DataSet();
        string sCompany = "15";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = "REGION_CODE";
        string OrderDir = "ASC";
        bool bNet = false;
        bool bRegion = false;
        string RateClass = "RATE_CLASS";
        string CONTRACT_CODE = "T";
        string CONTRACT_REASON = "SR";
        int nRowct = 0;
         string sType = "type";
         Int32[] arr_NB = new Int32[] { };
         string[] strRowValue = new string[3];
         DataTable datatab = new DataTable(); // Create a new Data table
         string RegionCodeAll = "ALL";
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
            if (!IsPostBack)
            {
                
                this.rdListType.SelectedIndexChanged += new EventHandler(rdListType_SelectedIndexChanged);
                tblgrid.Visible = false;
                txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                
                if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                    AgentID = Session["LoginID"].ToString();

                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();
                

                
               
            }

            //rdListType.Items[0].Selected = true;
            if (IsPostBack)
            {

                //rdListType.ClearSelection();
            }
        }

        void rdListType_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            throw new NotImplementedException();
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



            string sDuration = " ";
            //string[] fromDate = txtFrom.Text.Split('/');
            //FromDate = fromDate[2] + fromDate[0] + fromDate[1];

            //string[] toDate = txtTo.Text.Split('/');
            //ToDate = toDate[2] + toDate[0] + toDate[1];

            con.Open();
            //Calculating the duration of lapses
            string bType = "ALL";
            if(rdListType.Items[0].Selected)
            {
                bType = "ALL";
            }
            else if (rdListType.Items[1].Selected)
            {
                bType = "FirstYear";
                sDuration = " AND (DURATION>=1 AND DURATION<=12) ";
            }
            else if (rdListType.Items[2].Selected)
            {
                bType = "SecondYear";
                sDuration = " AND (DURATION>=13 AND DURATION<=24) ";
            }
            else if(rdListType.Items[3].Selected)
            {
                bType = "Renewal";
                sDuration = " AND (DURATION>24) ";
            }
            SqlCommand commPolicy = new SqlCommand();

            commPolicy.Connection = con;
            //commPolicy.CommandType = CommandType.StoredProcedure;
            commPolicy.CommandText = "SELECT REGION_CODE, REGION_NAME, (select count(distinct AGENT_NUMBER) FROM LAPSE L WHERE (LAPSE = 1) AND "+
                "(L.COMPANY_CODE = '"+ sCompany + "') AND AGENT_NUMBER in (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE (COMPANY_CODE = '15') AND "+
            " HIERARCHY_AGENT = '" + AgentID + "') and L.REGION_CODE = LL.REGION_CODE AND(LAST_CHANGE_DATE between '" + FromDate + "' AND '" + ToDate + "' )) as AGENTS_LAPSE, sum(case when((LAPSE = 1)AND (LAST_CHANGE_DATE between '"+FromDate+"' AND '"+ToDate+"' ))then 1 else 0 end) as LAPSE,"+ 
" sum(case when((LAPSE <> 3)AND(LAST_CHANGE_DATE between '"+FromDate+"' AND '"+ToDate+"'))then 1 else 0 end) as PAID,"+
 " sum(case when(INFORCE = 1) then 1 else 0 end) as INFORCE, sum(case when((RETURNED_CHECK = 1)AND(LAPSE = 1)AND"+
  " (LAST_CHANGE_DATE between '"+FromDate+"' AND '"+ToDate+"' ))then 1 else 0 end) as RETURNED_CHECK,"+
   " sum(case when((REPLACEMENT <> 0)AND(LAPSE = 1)AND(LAST_CHANGE_DATE between '"+FromDate+"' AND '"+ToDate+"' ))then 1 else 0 end) as REPLACEMENT,"+
    " sum(case when((RESCISSION <> 0)AND(LAPSE = 1)AND(LAST_CHANGE_DATE between '"+FromDate+"' AND '"+ToDate+"' ))then 1 else 0 end) as RESCISSIONS,"+
     " sum(case when((RTRIM(RATE_CLASS) = '"+ RateClass + "')AND(LAPSE=1)AND(LAST_CHANGE_DATE between '"+FromDate+"' AND '"+ToDate+"' ))then 1 else 0 end)"+
      " as TOBACCO_USE, sum(case when((REPLACEMENT = 2)AND(LAST_CHANGE_DATE between '"+FromDate+"' AND '"+ToDate+"' ))then 1 else 0 end) as W_REPLACEMENT,"+
       " sum(case when((REPLACEMENT = 1)AND(LAST_CHANGE_DATE between  '" + FromDate + "' AND '" + ToDate + "' ))then 1 else 0 end) as O_REPLACEMENT," +
        " sum(case when(((CONTRACT_CODE='T')AND(CONTRACT_REASON='" + CONTRACT_REASON+ "'))AND(LAST_CHANGE_DATE between  '" + FromDate + "' AND '" + ToDate + "' ))then 1 else 0 end)" +
         "as SURRENDERED, sum(case when((BUS_AT_RISK = 1))then 1 else 0 end) as BUS_AT_RISK FROM LAPSE LL WITH (NOLOCK) WHERE (LL.COMPANY_CODE =  '"+ sCompany + "')"+
          "AND LL.AGENT_NUMBER in (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WITH (NOLOCK) WHERE (COMPANY_CODE = LL.COMPANY_CODE)AND"+
          "(HIERARCHY_AGENT = '" + AgentID + "')) AND REGION_NAME <> '' " + sDuration + " GROUP BY REGION_CODE, REGION_NAME ORDER BY RETURNED_CHECK DESC";

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

                dvgrid.Visible = true;
                lblcount.Visible = true;
                //Added newly to implement Export to Excel functionality
                if (dataPolicy != null && dataPolicy.DefaultView != null)
                    dataPolicy = dataPolicy.DefaultView.ToTable();

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
            Export("LapsesByRegion.xls", this.grRequirementsByRegion);
        }

        
        
        protected void ExportToExcel_OLD()
        {
                InvokeSP();
             
                if (datatab.Rows.Count > 0 && datatab != null)
            {                
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"LapsesByRegion" + DateTime.Now.ToString();                
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
