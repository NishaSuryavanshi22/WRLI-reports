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
    public partial class Report3P : System.Web.UI.Page
    {
        SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //string agent = "WRE";
        DataSet dsPolicy = new DataSet();
        
        DataTable datatab = new DataTable(); // Create a new Data table
        public static DataTable dataPolicy = new DataTable();

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
         //DataTable datatab = new DataTable(); // Create a new Data table
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

            if (!IsPostBack)
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

                SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + AgentID + "' ORDER BY SORTNAME ASC");
                //Modified by siva
                //SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,DISPLAYNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + AgentID + "' AND AGENT_NUMBER = '" +
                    //AgentID + "' ORDER BY SORTNAME ASC");
                
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
        }

        protected void InvokeSP()
        {

            string[] fromDate = txtFrom.Text.Split('/');
            FromDate = fromDate[2] + fromDate[0] + fromDate[1];

            string[] toDate = txtTo.Text.Split('/');
            ToDate = toDate[2] + toDate[0] + toDate[1];

            con.Open();
            string bType = "ALL";
            string sAgentList = "";
            SqlCommand commPolicy = new SqlCommand();
            sAgentList = ddlAgentList.SelectedItem.Value;
            string[] SelAgentID = sAgentList.Split('-');

            if (SelAgentID != null && SelAgentID.Length > 0)
            {
                AgentID = SelAgentID[0].ToString();
            }
            
            commPolicy.Connection = con;
            commPolicy.Parameters.Add("@agentid",SqlDbType.VarChar,16);
            commPolicy.Parameters["@agentid"].Value = AgentID.Trim();
            //
            commPolicy.Parameters.Add("@company", SqlDbType.VarChar, 16);
            commPolicy.Parameters["@company"].Value = sCompany.Trim();
            //
            commPolicy.Parameters.Add("@fromdate", SqlDbType.VarChar, 16);
            commPolicy.Parameters["@fromdate"].Value = FromDate.Trim();
            //
            commPolicy.Parameters.Add("@todate", SqlDbType.VarChar, 16);
            commPolicy.Parameters["@todate"].Value = ToDate.Trim();
            //
            commPolicy.Parameters.Add("@orderby", SqlDbType.VarChar, 16);
            commPolicy.Parameters["@orderby"].Value = Orderby.Trim();
            //
            commPolicy.Parameters.Add("@orderdir", SqlDbType.VarChar, 16);
            commPolicy.Parameters["@orderdir"].Value = OrderDir.Trim();

            commPolicy.CommandType = CommandType.StoredProcedure;
            //commPolicy.CommandText = "dbo.AGENT_3P_EX_Test";
            commPolicy.CommandText = "dbo.AGENT_3P_EX";


            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);
            dataadapter.SelectCommand = commPolicy;
            //dataadapter.SelectCommand = commPolicy;
           // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // Create a SQL Data Adapter and assign it the cmd value. 
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

            grRequirementsByRegion.DataSource = datatab;
            grRequirementsByRegion.DataBind();

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
            grRequirementsByRegion.DataSource = MyDataView1;
            //DataRowCurrView.EndEdit();
            grRequirementsByRegion.DataSource = datatab;
            grRequirementsByRegion.DataBind();

        }


        protected void ddlagent_PreRender(object sender, EventArgs e)
        {
            if (ddlAgentList.Items.FindByValue(string.Empty) == null)
            {
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
            //Export("Report3P.xls", this.grRequirementsByRegion);
        }

                
        protected void ExportToExcel()
        {
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                dataPolicy.DefaultView.Sort = "REGION_CODE";
                dataPolicy = dataPolicy.DefaultView.ToTable();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"Reports3P" + DateTime.Now.ToString();                
                Response.AddHeader("Content-Disposition", "inline;filename=" + filename.Replace("/", "").Replace(":", "")+".xlsx");
                    //Call  Export function
                //Response.BinaryWrite(ExportToCSVFileOpenXML(datatab));   
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
