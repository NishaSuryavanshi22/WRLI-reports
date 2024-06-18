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
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System.Configuration;
using DocumentFormat.OpenXml.VariantTypes;
using System.Web.Services.Description;

namespace WRLI_Reports
{
    public partial class ProductReport_By_Agent : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        // SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //string agent = "WRE";
        //    SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        DataSet dsPolicy = new DataSet();
        
        DataTable dataPolicy = new DataTable();
        //string agent = "WRE";
        Int32[] arr_NB = new Int32[] { };
        string sCompany = "15";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = "Agent_number";
        string OrderDir = "ASC";
        bool bNet = false;
        bool bRegion = false;
        int nRowct = 0;
         string sType = "type";
         //Int32[] arr_NB = new Int32[] { };
         string[] strRowValue = new string[3];
         DataTable datatab = new DataTable(); // Create a new Data table
         string RegionCodeAll = "ALL";
        string products = "";
        string Datatype = "";


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
            //ddlregion.Items.Clear();
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

                //Bind Agent
                SqlCommand comm = new SqlCommand("SELECT DISTINCT AGENT_NUMBER,DISPLAYNAME,SORTNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + AgentID + "' ORDER BY AGENT_NUMBER ASC");
                //Bind Agent new one
                               comm.Connection = con;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(comm.CommandText, con);
                ad.Fill(ds);
                // Data binding
                if (ds != null)
                    ds.Tables[0].DefaultView.Sort = "AGENT_NUMBER ASC";
                ddlregion.Items.Clear();
                ddlregion.DataSource = ds;
                ddlregion.DataTextField = "AGENT_NUMBER";
                //ddlregion.DataValueField = "AGENT_NUMBER";
                //Newly added siva
                ddlregion.DataValueField = "AGENT_NUMBER";
                ddlregion.DataBind();


                if (ddlregion.Items.FindByValue("ALL") == null)
                {
                    ddlregion.Items.Insert(0, new ListItem("ALL", "ALL"));
                        ddlregion.SelectedIndex = 0;
                    
                }
                //To get the Product from coverage table
                SqlCommand comnd = new SqlCommand("select product from coverage2");
                comnd.Connection = con;
                DataSet datas = new DataSet();
                SqlDataAdapter adpt = new SqlDataAdapter(comnd.CommandText, con);
                adpt.Fill(datas);
                //
               
                ddlProductTypes.Items.Clear();
                ddlProductTypes.DataSource = datas;
                ddlProductTypes.DataTextField = "product";
                //Newly added siva
                ddlProductTypes.DataValueField = "product";
                ddlProductTypes.DataBind();
                if (datas != null)
                    datas.Tables[0].DefaultView.Sort = "product ASC";
                if (ddlProductTypes.Items.FindByValue("ALL") == null)
                {
                    ddlProductTypes.Items.Insert(0, new ListItem("ALL", "ALL"));
                    ddlProductTypes.SelectedIndex = 0;

                }

                //nisha
                //To get the Region code form policies table
                SqlCommand comn = new SqlCommand("SELECT distinct REGION_CODE FROM POLICIES2 where REGION_CODE <> '' ");
                comnd.Connection = con;
                DataSet dataset = new DataSet();
                SqlDataAdapter dtadpt = new SqlDataAdapter(comn.CommandText, con);
                dtadpt.Fill(dataset);
                //
               
                ddlregioncode.Items.Clear();
                ddlregioncode.DataSource = dataset;
                ddlregioncode.DataTextField = "region_code";
                ddlregioncode.DataValueField = "region_code";
                ddlregioncode.DataBind();
                if (dataset != null)
                    dataset.Tables[0].DefaultView.Sort = "region_code ASC";
                if (ddlregioncode.Items.FindByValue("ALL") == null)
                {
                    ddlregioncode.Items.Insert(0, new ListItem("ALL", "ALL"));
                    ddlregioncode.SelectedIndex = 0;

                }

                if (Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "")
                {
                    if (Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "")
                        FromDate = Request.QueryString["fromdate"].ToString();
                    if (Request.QueryString["Todate"] != null && Request.QueryString[""] != "")
                        ToDate = Request.QueryString["Todate"].ToString();
                    if (Request.QueryString["Region_code"] != null && Request.QueryString["Region_code"] != "")
                        sRegionCode = Request.QueryString["Region_code"].ToString();
                    if (Request.QueryString["Company_code"] != null && Request.QueryString["Company_code"] != "")
                        sCompany = Request.QueryString["Company_code"].ToString();
                    if (Request.QueryString["Datatype"] != null && Request.QueryString["Datatype"] != "")
                        Datatype = Request.QueryString["Datatype"].ToString();
                    if (Request.QueryString["Products"] != null && Request.QueryString["Products"] != "")
                        products = Request.QueryString["Products"].ToString();


                    InvokeSP();
                }
            }
            con.Close();

        }

        protected void InvokeSP()
        {
            string[] fromDate;
            string[] toDate;
            if (Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "")
            {
                if (Request.QueryString["fromdate"] != null && Request.QueryString["fromdate"] != "")
                    FromDate = Request.QueryString["fromdate"].ToString();
                if (Request.QueryString["Todate"] != null && Request.QueryString[""] != "")
                    ToDate = Request.QueryString["Todate"].ToString();
                if (Request.QueryString["Region_code"] != null && Request.QueryString["Region_code"] != "")
                    sRegionCode = Request.QueryString["Region_code"].ToString();
                if (Request.QueryString["Company_code"] != null && Request.QueryString["Company_code"] != "")
                    sCompany = Request.QueryString["Company_code"].ToString();
                if (Request.QueryString["Datatype"] != null && Request.QueryString["Datatype"] != "")
                    Datatype = Request.QueryString["Datatype"].ToString();
                if (Request.QueryString["Products"] != null && Request.QueryString["Products"] != "")
                    products = Request.QueryString["Products"].ToString();

            }
            else
            {
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

                if (ddlProductTypes.SelectedValue != null)
                    products = ddlProductTypes.SelectedValue.ToString();

                if (ddlregioncode.SelectedValue != null)
                    sRegionCode = ddlregioncode.SelectedValue.ToString();

                if (rdListType.SelectedValue != null)
                    Datatype = rdListType.SelectedValue.ToString();
                con.Open();
                //string bType = "ALL";
                string sAgentList = "";
                if (ddlregion == null || ddlregion.SelectedItem == null)
                    return;
                sAgentList = ddlregion.SelectedItem.Value;
                string[] SelAgentID = sAgentList.Split('-');


                if (SelAgentID != null && SelAgentID.Length > 0)
                {
                    AgentID = SelAgentID[0].ToString();
                }
            }

            SqlCommand commPolicy = new SqlCommand();
            commPolicy.Connection = con;

            //            commPolicy.CommandText = "select pr.REGION_CODE as region_code, dbo.GETREGIONNAME(pr.region_code) as REGION_NAME, "+ 
            //"sum(case when(pr.contract_code='"+ "A" +"') then (CASE WHEN (ST.PROD_PCNT IS NULL) THEN 1 "+
            //"ELSE (ST.PROD_PCNT/100)end) else 0 end) as inforce," +
            //"sum(case when(pr.contract_code='"+ "P" +"') then (CASE WHEN (ST.PROD_PCNT IS NULL) THEN 1 "+
            //"ELSE (ST.PROD_PCNT/100)end) else 0 end) as pending, sum(case when("+
            //"((PR.APP_RECEIVED_DATE BETWEEN '"+ FromDate+"' AND "+  
            //"'"+  ToDate +"')))then (CASE WHEN (ST.PROD_PCNT IS NULL) THEN 1 "+
            //"ELSE (ST.PROD_PCNT/100) END)  else 0 end) as count, "+ 		
            //"sum( case when (ST.PROD_PCNT IS NULL) then annual_premium "+
            //" else annual_premium*(ST.PROD_PCNT/100)end) as annualizedprem, "+
            //"sum( case when (ST.PROD_PCNT IS NULL) then face_amount "+
            //" else face_amount*(ST.PROD_PCNT/100)end) as face_amount from policies2 pr left OUTER JOIN policy_split st ON  "+
            //"(PR.COMPANY_CODE = ST.COMPANY_CODE) and (PR.POLICY_NUMBER = ST.POLICY_NUMBER) AND "+
            //"(PR.AGENT_NUMBER = ST.AGENT_NUMBER) and (PR.ISSUE_DATE = ST.ISSUE_DATE)  "+ 		
            //"WHERE (pr.COMPANY_CODE='" + sCompany + "') AND pr.AGENT_NUMBER in  "+  
            //"(SELECT AH.AGENT_NUMBER FROM AGENT_HIERLIST AH WHERE "+  
            //"(AH.COMPANY_CODE='" + sCompany + "') "+
            //"AND AH.HIERARCHY_AGENT = '"+AgentID+"') and ( pr.app_received_date between"+
            //"'"+ FromDate +"' AND '"+ ToDate +"') and product_code in(select coverage_id from coverage2 where product in('"+ products +"')) "+
            //"and ((pr.region_code='"+  ddlregion.SelectedValue +"'))" +	
            //"GROUP BY pr.region_code ORDER BY " + Orderby + "";

            //            commPolicy.CommandText = "select pr.REGION_CODE as region_code, dbo.GETREGIONNAME(pr.region_code) as REGION_NAME, " +
            //"sum(case when(pr.contract_code='" + "A" + "') then (CASE WHEN (ST.PROD_PCNT IS NULL) THEN 1 " +
            //"ELSE (ST.PROD_PCNT/100)end) else 0 end) as inforce," +
            //"sum(case when(pr.contract_code='" + "P" + "') then (CASE WHEN (ST.PROD_PCNT IS NULL) THEN 1 " +
            //"ELSE (ST.PROD_PCNT/100)end) else 0 end) as pending, sum(case when(" +
            //"((PR.APP_RECEIVED_DATE BETWEEN '" + FromDate + "' AND " +
            //"'" + ToDate + "')))then (CASE WHEN (ST.PROD_PCNT IS NULL) THEN 1 " +
            //"ELSE (ST.PROD_PCNT/100) END)  else 0 end) as count, " +
            //"sum( case when (ST.PROD_PCNT IS NULL) then annual_premium " +
            //" else annual_premium*(ST.PROD_PCNT/100)end) as annualizedprem, " +
            //"sum( case when (ST.PROD_PCNT IS NULL) then face_amount " +
            //" else face_amount*(ST.PROD_PCNT/100)end) as face_amount from policies2 pr left OUTER JOIN policy_split st ON  " +
            //"(PR.COMPANY_CODE = ST.COMPANY_CODE) and (PR.POLICY_NUMBER = ST.POLICY_NUMBER) AND " +
            //"(PR.AGENT_NUMBER = ST.AGENT_NUMBER) and (PR.ISSUE_DATE = ST.ISSUE_DATE)  " +
            //"WHERE (pr.COMPANY_CODE='" + sCompany + "') AND pr.AGENT_NUMBER in  " +
            //"(SELECT AH.AGENT_NUMBER FROM AGENT_HIERLIST AH WHERE " +
            //"(AH.COMPANY_CODE='" + sCompany + "') " +
            //"AND AH.HIERARCHY_AGENT = '" + AgentID + "') and ( pr.app_received_date between" +
            //"'" + FromDate + "' AND '" + ToDate + "') and product_code in(select coverage_id from coverage2 where product in('" + products + "')) "  +
            //"GROUP BY pr.region_code ORDER BY " + Orderby + "";

            //nisha
            if (Datatype == "Submitted")
            {
                commPolicy.CommandText = "select pr.agent_number as agent_number," +
                    "dbo.GET_AGENT_DISPLAY_NAME(pr.company_code,pr.agent_number,'L') as AGENT_NAME, " +
"sum(case when(pr.contract_code='" + "A" + "') then (CASE WHEN (ST.PROD_PCNT IS NULL) THEN 1 " +
"ELSE (ST.PROD_PCNT/100)end) else 0 end) as inforce," +
"sum(case when(pr.contract_code='" + "P" + "') then (CASE WHEN (ST.PROD_PCNT IS NULL) THEN 1 " +
"ELSE (ST.PROD_PCNT/100)end) else 0 end) as pending," +
"sum(case when(((PR.ISSUE_DATE BETWEEN '" + FromDate + "' " +
"AND '" + ToDate + "')AND(PR.RECORD_TYPE='I'))" +
"or((PR.APP_RECEIVED_DATE BETWEEN '" + FromDate + "' AND " +
"'" + ToDate + "')AND(PR.RECORD_TYPE='P')))then (CASE WHEN (ST.PROD_PCNT IS NULL) " +
"THEN 1 " +
"ELSE (ST.PROD_PCNT/100) " +
"END)  else 0 end) as count, " +


"sum( case when (ST.PROD_PCNT IS NULL) then annual_premium " +

" else annual_premium*(ST.PROD_PCNT/100)end) as annualizedprem, " +
"sum( case when (ST.PROD_PCNT IS NULL) then face_amount " +

" else face_amount*(ST.PROD_PCNT/100)end) as face_amount " +
"from policies2 pr " +
"left OUTER JOIN policy_split st ON  " +
"(PR.COMPANY_CODE = ST.COMPANY_CODE) and (PR.POLICY_NUMBER = ST.POLICY_NUMBER) AND " +
"(PR.AGENT_NUMBER = ST.AGENT_NUMBER) and (PR.ISSUE_DATE = ST.ISSUE_DATE)  " +


"WHERE (pr.COMPANY_CODE='" + sCompany + "') AND pr.AGENT_NUMBER in  " +
"(SELECT AH.AGENT_NUMBER FROM AGENT_HIERLIST AH WHERE " +
"(AH.COMPANY_CODE='" + sCompany + "') " +
"AND AH.HIERARCHY_AGENT = '" + AgentID + "') and ( pr.app_received_date between" +
"'" + FromDate + "' AND '" + ToDate + "') and ((pr.region_code='" + sRegionCode + "')) and product_code in(select coverage_id from coverage2 where product in('" + products + "')) " +
"GROUP BY pr.company_code, pr.agent_number ORDER BY " + Orderby + "";

            }
            else {
                commPolicy.CommandText = "  select pr.agent_number as agent_number,  " +  		
"dbo.GET_AGENT_DISPLAY_NAME(pr.company_code,pr.agent_number,'L') as AGENT_NAME, " + 

"sum(case when(pr.contract_code='" + "A "+ "') then (CASE WHEN (ST.PROD_PCNT IS NULL) " + 

"THEN 1 " + 
"ELSE (ST.PROD_PCNT/100)end) " + 
"else 0 end) as inforce," + 
"sum(CASE WHEN ((PR.ISSUE_DATE BETWEEN '" + FromDate + "' AND '" + ToDate + "') AND " + 
 "(PR.PAYMENT_FLAG =' " +" Y" + "')) then (CASE WHEN (ST.PROD_PCNT IS NULL) " + 
"then 1 " + 
"else (ST.PROD_PCNT/100) " + 
"end)  else 0 end) as count, " + 
"sum( case when (ST.PROD_PCNT IS NULL) then annual_premium " + 
" else annual_premium*(ST.PROD_PCNT/100)end) as annualizedprem, " + 
"sum( case when (ST.PROD_PCNT IS NULL) then face_amount " + 
" else face_amount*(ST.PROD_PCNT/100)end) as face_amount " +
"from policies2 pr " + 
"left OUTER JOIN policy_split st ON  " + 
"(PR.COMPANY_CODE = ST.COMPANY_CODE) and (PR.POLICY_NUMBER = ST.POLICY_NUMBER) AND " + 
"(PR.AGENT_NUMBER = ST.AGENT_NUMBER) and (PR.ISSUE_DATE = ST.ISSUE_DATE)  " +
"WHERE (pr.COMPANY_CODE='" + sCompany + "') " + 
"and ( " + 
"(pr.ISSUE_DATE BETWEEN '" + FromDate + "' AND '" + ToDate + "')) " +
"and ((pr.region_code='" + ddlregioncode.SelectedValue + "')) and "+
"pr.AGENT_NUMBER in (SELECT AH.AGENT_NUMBER FROM AGENT_HIERLIST AH WHERE (AH.COMPANY_CODE='" + sCompany + "')  "+"AND AH.HIERARCHY_AGENT = '" + AgentID + "') "+
"and product_code in(select coverage_id from coverage2 where product in('" + products + "')) " +
"GROUP BY pr.agent_number,pr.company_code ORDER BY " + Orderby + "";
            }
            
              
            //Siva
            //Binding agent/region new
            /* SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,ISNULL(DISPLAYNAME,'') AS DISPLAYNAME,SORTNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT = '" + AgentID + "' ORDER BY AGENT_NUMBER ASC");*/

            SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy.CommandText, con);
            dataadapter.SelectCommand = commPolicy;
            //dataadapter.SelectCommand = commPolicy;
           // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // C'reate a SQL Data Adapter and assign it the cmd value. 
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

            //con.Close();
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
            if (ddlregion.Items.FindByValue(string.Empty) == null)
            {
                if(! IsPostBack)
                    ddlregion.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
        }

        protected void ddlregion_PreRender(object sender, EventArgs e)
        {
            
        }

        protected void ddlProductTypes_PreRender(object sender, EventArgs e)
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
            //Export("PaidBusiness.xls", this.grdSubmittedReport);
        }



        protected void ExportToExcel()
        {
            //InvokeSP();
            //            dataPolicy = Session["dataPolicy"] as DataTable;
            if (datatab.Columns.Count <= 0 )
                return;

//            datatab.DefaultView.Sort = "REGION_CODE";
            datatab = datatab.DefaultView.ToTable();
            //dataPolicy = GridView1.DataSource as DataTable;
            if (datatab.Rows.Count > 0 && datatab != null)
            {
                datatab.DefaultView.Sort = "REGION_CODE";
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
