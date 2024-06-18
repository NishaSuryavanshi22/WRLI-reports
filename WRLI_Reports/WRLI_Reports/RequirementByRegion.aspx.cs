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
    public partial class RequirementsByRegion : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        // SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());

        //  SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        string agent = "WRE";
        DataTable datatab = new DataTable(); // Create a new Data table
        public static DataTable dataPolicy = new DataTable();

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
                    sCompany = "07";
                if (Session["LoginID"] != null && Session["LoginID"].ToString() != "")
                    AgentID = Session["LoginID"].ToString();

                if (Session["RegionCode"] != null && Session["RegionCode"].ToString() != "")
                    sRegionCode = Session["RegionCode"].ToString();

                //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Initial Catalog=WRE_AGENT");
               
            }
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

            //string[] fromDate = txtFrom.Text.Split('/');
            //FromDate = fromDate[2] + fromDate[0] + fromDate[1];

            //string[] toDate = txtTo.Text.Split('/');
            //ToDate = toDate[2] + toDate[0] + toDate[1];

            con.Open();
            string bType = "ALL";
            if(rdListType.Items[0].Selected)
            {
                bType = "ALL";
            }
            else if (rdListType.Items[2].Selected)
            {
                bType = "NON-MED";
            }
            else if (rdListType.Items[1].Selected)
            {
                bType = "MED";
            }
            else if(rdListType.Items[3].Selected)
            {
                bType = "MortgageCertificate";
            }

            SqlCommand commPolicy = new SqlCommand();

            commPolicy.Connection = con;
            //commPolicy.CommandType = CommandType.StoredProcedure;
            commPolicy.CommandText = " select (UR.REGION_CODE), ISNULL(REGION_NAME,'') AS REGION_NAME, sum(case when (((PB.SL_PERCENT > 0) AND (PB.SL_PCT_CEASE_DATE>'20160606'))OR ((PB.SL_FLAT_AMOUNT > 0) AND"+ " (PB.SL_FLAT_CEASE_DATE>'20160604'))) THEN 1 ELSE 0 END) AS RATING, cast((cast(sum(case when (((PB.SL_PERCENT > 0) AND (PB.SL_PCT_CEASE_DATE>'20160604'))OR "+
"((PB.SL_FLAT_AMOUNT > 0) AND (PB.SL_FLAT_CEASE_DATE>'20160604'))) THEN 1 ELSE 0 END) as decimal(9,2))/count(distinct(UR.POLICY_NUMBER)))*100 as decimal(9,2)) as RATING_PERCENT,"+
    " sum(PROD_PCNT)as APP_COUNT, SUM(APS) as APS_COUNT, cast((cast(SUM(APS) as decimal(9,2))/count(distinct(UR.POLICY_NUMBER)))*100 as decimal(9,2)) as APS_PERCENT,"+
    " SUM(PARAMED) as PARAMED_COUNT, cast((cast(SUM(PARAMED) as decimal(9,2))/count(distinct(UR.POLICY_NUMBER)))*100 as decimal(9,2)) as PARAMED_PERCENT, SUM(HOS) as HOS_COUNT, "+
    " cast((cast(SUM(HOS) as decimal(9,2))/count(distinct(UR.POLICY_NUMBER)))*100 as decimal(9,2)) as HOS_PERCENT, SUM(BLOOD) as BLOOD_COUNT, cast((cast(SUM(BLOOD) "+
            " as decimal(9,2))/count(distinct(UR.POLICY_NUMBER)))*100 as decimal(9,2)) as BLOOD_PERCENT, SUM(IBU) as IBU_COUNT, cast((cast(SUM(IBU) as "+
            " decimal(9,2))/count(distinct(UR.POLICY_NUMBER)))*100 as decimal(9,2)) as IBU_PERCENT, SUM(PHI) as PHI_COUNT, cast((cast(SUM(PHI) as "+
            " decimal(9,2))/count(distinct(UR.POLICY_NUMBER)))*100 as decimal(9,2)) as PHI_PERCENT, SUM(GIS) as GIS_COUNT, cast((cast(SUM(GIS) as "+
            " decimal(9,2))/count(distinct(UR.POLICY_NUMBER)))*100 as decimal(9,2)) as GIS_PERCENT, SUM(AMEND) as AMEND_COUNT, cast((cast(SUM(AMEND) as "+
            " decimal(9,2))/count(distinct(UR.POLICY_NUMBER)))*100 as decimal(9,2)) as AMEND_PERCENT from U_REQ UR inner join POLICIES2 PO ON "+
            " PO.POLICY_NUMBER = UR.POLICY_NUMBER and PO.COMPANY_CODE = UR.COMPANY_CODE left outer join POLICY_COVERAGE PB ON PB.POLICY_NUMBER = UR.POLICY_NUMBER "+
            " and PB.COMPANY_CODE = UR.COMPANY_CODE AND (PB.BENEFIT_TYPE = 'SL') WHERE UR.AGENT_NUMBER in (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = '"+AgentID +"') "+
            " AND UR.COMPANY_CODE = '"+  sCompany +"' AND UR.REGION_CODE IS NOT NULL AND (UR.APP_RECEIVED_DATE BETWEEN '"+ FromDate +"' AND '"+ToDate  +"') AND ((LEFT(UR.REGION_CODE,1) = '"+RegionCodeAll+ "')"+
            " OR (SUBSTRING(UR.REGION_CODE,2,1) = '" + RegionCodeAll + "' ) OR (LEFT(UR.REGION_CODE,2) = '" + RegionCodeAll + "') OR ('ALL' = '" + RegionCodeAll + "')) AND ((MED_TYPE = '" + bType + "') OR " +
            " ('" + bType + "' = 'ALL') OR  (MED_TYPE = 'NON-MED' and MED_TYPE = '" + bType + "' and PO.FACE_AMOUNT > 100000 and PO.FACE_AMOUNT <= 150000)) GROUP BY UR.REGION_CODE,REGION_NAME ORDER BY UR.REGION_CODE ";
 

            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);
            
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
                AddEditRows();
                lblcount.Text = "Total Policy Count: " + datatab.Rows.Count.ToString();
                if (dataPolicy != null && dataPolicy.DefaultView != null)
                    dataPolicy = dataPolicy.DefaultView.ToTable();
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
            if (DataRowCurrView["REGION_NAME"] != null && DataRowCurrView["REGION_NAME"].ToString().Trim() != "")
                strRowValue[0] +=DataRowCurrView["REGION_NAME"] + "~";
            if (DataRowCurrView["REGION_CODE"] != null && DataRowCurrView["REGION_CODE"].ToString().Trim() != "")
                strRowValue[1] += DataRowCurrView["REGION_CODE"] + "~";
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
