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

//using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace WRLI_Reports
{
    public partial class SubmittedCount : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //  SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 1000");
        string agent = "WRE";
        public static DataTable dataPolicy = new DataTable();
        DataSet dsPolicy = new DataSet();
        
        string sCompany = "15";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = "PAY_MONTH,PR.REGION_CODE";
        string OrderDir = "ASC";
        bool bNet = false;
        bool bRegion = false;
        int nRowct = 0;
         string sType = "type";
         Int32[] arr_NB = new Int32[] { };
         DataTable datatab = new DataTable(); // Create a new Data table

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
            //agent = Session["LoginID"].ToString();
            if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                sCompany = Session["CompanyCode"].ToString();
            else
                sCompany = "15";

            if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                AgentID = Session["LoginID"].ToString();

            if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                sRegionCode = Session["RegionCode"].ToString();
            Label1.Text = "Region: " + sRegionCode + ", Agent: " + AgentID + ", Company: " + sCompany;
            if (!IsPostBack)
            {
                tblgrid.Visible = false;
                txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                

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

            SqlCommand commPolicy = new SqlCommand();

            commPolicy.Connection = con;
            commPolicy.CommandType = CommandType.StoredProcedure;
            commPolicy.CommandText = "NET_BY_REGION_Test";
            SqlDataAdapter adPolicy = new SqlDataAdapter(commPolicy.CommandText, con);
            //commPolicy.Parameters.Add("@RETURN_VALUE", SqlDbType.VarChar).Value = "1";
            commPolicy.Parameters.AddWithValue("@agentid", AgentID);
            if (!bNet && bRegion)
                commPolicy.Parameters.AddWithValue("@regioncode", SqlDbType.VarChar).Value = sRegionCode;
            commPolicy.Parameters.AddWithValue("@company", SqlDbType.VarChar).Value = sCompany;
            commPolicy.Parameters.AddWithValue("@fromdate", SqlDbType.VarChar).Value = FromDate;
            commPolicy.Parameters.AddWithValue("@todate", SqlDbType.VarChar).Value = ToDate;
            commPolicy.Parameters.AddWithValue("@orderby", SqlDbType.VarChar).Value = Orderby;
            commPolicy.Parameters.AddWithValue("@orderdir", SqlDbType.VarChar).Value = OrderDir;


            if (bNet)
            {
                if (bRegion)
                {
                    commPolicy.CommandText = "NET_BY_REGION_PER_AGENT";
                }
                else
                {
                    commPolicy.CommandText = "NET_BY_REGION";
                }

            }
            else if (bRegion)
            {
                commPolicy.CommandText = "SUBMITTED_BY_REGION_PER_AGENT";
            }
            else
            {
                //commPolicy.CommandText = "SUBMITTED_BY_REGION_Test";
                commPolicy.CommandText = "SUBMITTED_BY_REGION";
            }

            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // Create a SQL Data Adapter and assign it the cmd value. 
            datatab = new DataTable(); // Create a new Data table
            dataadapter.Fill(datatab);
            //Fill the data table for export excel
            if (datatab != null)
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
            arr_NB[4] = new Int32();

        }
        protected void ReadColumnValues(DataRowView DataRowCurrView, ref Int32[] arr_NB)
        {
            //arr_NB[0] += Convert.ToInt32(DataRowCurrView["DISPLAYID"]);
            //arr_NB[1] += Convert.ToInt32(DataRowCurrView["DISPLAYNAME"]);
            if (DataRowCurrView["SUB_COUNT"] != null && DataRowCurrView["SUB_COUNT"].ToString().Trim() != "")
                arr_NB[3] += Convert.ToInt32(DataRowCurrView["SUB_COUNT"]);
            if (DataRowCurrView["SUB_PREM"] != null && DataRowCurrView["SUB_PREM"].ToString().Trim() != "")
                arr_NB[4] += Convert.ToInt32(DataRowCurrView["SUB_PREM"]);
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

            //
            for (int nIndex = 0; nIndex < nRowct; nIndex++)
            {

                DataRowCurrView = MyDataView1[nIndex];
                //List<Int32> stringList = new List<Int32>();
                ReadColumnValues(DataRowCurrView, ref arr_NB);
                // End of new logic
            }
            //Below line Commented by Siva 07 Jan
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
                MyDataRowView["SUB_COUNT"] = arr_NB[3];
                MyDataRowView["SUB_PREM"] = arr_NB[4];
                
            }
            grdSubmitCount.DataSource = MyDataView1;
            grdSubmitCount.DataBind();

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
            //Export("PolicyReport.xls", this.grdSubmitCount);
            ExportToExcel();



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

        protected void ExportToExcel()
        {
            //InvokeSP();
            //            dataPolicy = Session["dataPolicy"] as DataTable;

            //dataPolicy = GridView1.DataSource as DataTable;
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                dataPolicy.DefaultView.Sort = "DISPLAYNAME";
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
       
    }

}
