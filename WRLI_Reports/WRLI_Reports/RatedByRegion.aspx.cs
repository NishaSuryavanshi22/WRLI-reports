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
    public partial class RatedByRegion : System.Web.UI.Page
    {
        //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        // SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //   SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        //string agent = "WRE";

        DataSet dsPolicy = new DataSet();
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
         //string sType = "type";
         Int32[] arr_NB = new Int32[] { };
         string[] strRowValue = new string[3];
         DataTable datatab = new DataTable(); // Create a new Data table
         DataTable datatabTotal = new DataTable();
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

                if (Request.QueryString["State"] != null && Request.QueryString["State"] != "")
                {
                    string sState = Request.QueryString["State"];
                    InvokeSP();
                   // AddEditRows();
                }



                //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Initial Catalog=WRE_AGENT");

            }

        }

        protected void InvokeSP()
        {
            //string sState = Request.QueryString["State"];


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

            string sPaidColumn ="PAYMENT_DATE";

            //string[] sFromDate = txtFrom.Text.Split('/');
            //FromDate = sFromDate[2] + sFromDate[0] + sFromDate[1];

            //string[] sToDate = txtTo.Text.Split('/');
            //ToDate = sToDate[2] + sToDate[0] + sToDate[1];

            con.Open();
            string bType = "ALL";
            if(rdListType.Items[0].Selected)
            {
                bType = "ALL";
            }
            SqlCommand commPolicy = new SqlCommand();
            string  sSubQuery = "((((PR."+ sPaidColumn + " BETWEEN " + FromDate+ " AND "+ ToDate+ ")AND(PR.RECORD_TYPE='I'))OR((PR.APP_RECEIVED_DATE BETWEEN "+
                FromDate+ " AND " + ToDate+ ")AND(PR.RECORD_TYPE='P'))) AND (PR.APP_RECEIVED_DATE IS NOT NULL))";

            string sPaidQuery = "((PR." + sPaidColumn + " BETWEEN " + FromDate + " AND " + ToDate + ") AND (PR.PAYMENT_FLAG = 'Y') )";
            string sNTOQuery = "((PR." + sPaidColumn + " BETWEEN " + FromDate + " AND " + ToDate + ") AND (PR.PAYMENT_FLAG = 'Y') AND ((PR.CONTRACT_CODE = 'T')AND(PR.CONTRACT_REASON = 'NT')))";

            commPolicy.Connection = con;
            //if (Request.QueryString["LatestRef"] != null && Request.QueryString["LatestRef"] != "")
            //Interview for agent and the else part is to get the latest ref ID 
            // If the check box is selected
            //if (chkInterview.Checked)
            
                commPolicy.CommandText = "select PR.REGION_CODE,ISNULL(RN.AGENCY_NAME,RN.AGENT_NAME) AS REGION_NAME,PR.COMPANY_CODE, SUM(CASE  WHEN ("+sSubQuery+")"+   " THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_COUNT, SUM(CASE  WHEN ("+sPaidQuery+")   THEN "+
                    " (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as PAID_COUNT, SUM(CASE  WHEN ("+sNTOQuery+")"+ " THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as NTO_COUNT,CASE WHEN (SUM(CASE  WHEN ("+sPaidQuery+")"+   " THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN ("+sPaidQuery+")THEN  "+
" (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN ISNULL(PR.ANNUAL_PREMIUM,0)    ELSE ISNULL(PR.ANNUAL_PREMIUM,0)*(ST.PROD_PCNT/100)  END)   ELSE 0 END)) /(SUM(CASE "+  " WHEN ("+sPaidQuery+")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END as AVG_PAID,CASE WHEN (SUM(CASE  "+
" WHEN ("+sSubQuery+")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  "+
"  WHEN("+sPaidQuery+")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN ("+sSubQuery+") "+  "THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END as PAID_PCNT,CASE WHEN (SUM(CASE  WHEN ("+sSubQuery+")"+ " THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN ("+sSubQuery+" AND "+  " PR.RATE_CLASS = 'Smoker')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  WHEN ("+sSubQuery+")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)  END)   ELSE 0   END)) END AS SUB_NIC_PCNT,CASE WHEN (SUM(CASE WHEN ("+sPaidQuery+")"+
" THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN ("+sPaidQuery+" AND "+  "PR.RATE_CLASS = 'Smoker')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE "+ 
            " WHEN ("+sPaidQuery+")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_NIC_PCNT,"+
"  SUM(CASE  WHEN ("+sSubQuery+" AND RATE_CLASS='Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100) "+ 
            " END)   ELSE 0   END) as SUB_NIC_L_COUNT,SUM(CASE  WHEN ("+sSubQuery+" AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'L') THEN  (CASE  WHEN (ST.PROD_PCNT "+
" IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_L_COUNT,SUM(CASE  WHEN ("+sSubQuery+" AND PLAN_CODE = 'L')   THEN   (CASE "+ 
" WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_L_COUNT,CASE WHEN (SUM(CASE  WHEN ("+sSubQuery+" AND "+
            " PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE "+
" (SUM(CASE  WHEN ("+sSubQuery+" AND  PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    "+
" ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  WHEN ("+sSubQuery+" AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL) THEN 1 ELSE "+
" (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_L_NIC_PCNT,CASE WHEN (SUM(CASE  WHEN ("+sPaidQuery+"AND PLAN_CODE = 'L')   THEN   (CASE  WHEN "+
" (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN ("+sPaidQuery+"AND   "+
" PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0 "+
" END))/(SUM(CASE  WHEN ("+sPaidQuery+"AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0 "+
" END)) end AS PAID_L_NIC_PCNT,SUM(CASE  WHEN ("+sSubQuery+" AND RATE_CLASS='Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1 "+   " ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NIC_G_COUNT,SUM(CASE  WHEN ("+sSubQuery+" AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  "+
" WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_G_COUNT,SUM(CASE  WHEN ("+sSubQuery+" AND PLAN_CODE = 'G')"+   " THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_G_COUNT,CASE WHEN (SUM(CASE  WHEN ("+sSubQuery+" AND"+
" PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN "+      " ("+sSubQuery+" AND  PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END) "+
" ELSE 0   END))  /(SUM(CASE  WHEN ("+sSubQuery+" AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END) "+
" ELSE 0   END)) END AS SUB_G_NIC_PCNT, CASE WHEN (SUM(CASE  WHEN ("+sPaidQuery+" AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1 "+
" ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN ("+sPaidQuery+" AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'G') "+
" THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN ("+sPaidQuery+"AND PLAN_CODE = 'G') "+
" THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_G_NIC_PCNT,SUM(CASE  WHEN ("+sSubQuery+" AND"+
" RATE_CLASS='Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as "+
" SUB_NIC_M_COUNT, SUM(CASE  WHEN ("+sSubQuery+" AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE "+
" (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_M_COUNT,SUM(CASE  WHEN ("+sSubQuery+" AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL) "+    " THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_M_COUNT, CASE WHEN (SUM(CASE  WHEN ("+sSubQuery+" AND PLAN_CODE = 'M')   THEN   (CASE  WHEN "+
" (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN ("+sSubQuery+" AND  "+
" PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0 "+
" END))  / (SUM(CASE  WHEN ("+sSubQuery+" AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END) "+
" ELSE 0   END)) END AS SUB_M_NIC_PCNT,CASE WHEN (SUM(CASE  WHEN ("+sPaidQuery+"AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1  "+
" ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN ("+sPaidQuery+"AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'M')   THEN "+
" (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/ (SUM(CASE  WHEN ("+sPaidQuery+"AND PLAN_CODE = 'M')   THEN "+
" (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_M_NIC_PCNT from POLICIES PR LEFT OUTER JOIN "+  " POLICY_SPLIT ST WITH (NOLOCK)  ON (PR.COMPANY_CODE = ST.COMPANY_CODE) and (PR.POLICY_NUMBER = ST.POLICY_NUMBER) AND  (PR.AGENT_NUMBER = ST.AGENT_NUMBER) and "+  " (PR.ISSUE_DATE = ST.ISSUE_DATE)  LEFT OUTER JOIN REGION_NAMES RN ON RN.MARKETING_COMPANY = PR.REGION_CODE where  PR.AGENT_NUMBER in (SELECT AHL0.AGENT_NUMBER "+
" FROM AGENT_HIERLIST AHL0 WHERE AHL0.COMPANY_CODE = '"+sCompany+"' AND AHL0.HIERARCHY_AGENT='"+AgentID +"') AND PR.COMPANY_CODE = '"+sCompany+"' AND  "+
" PR.RATE_CLASS IS NOT NULL AND PR.REGION_CODE <> ''  AND (('"+sState+"'='ALL')OR(PR.PI_STATE = '"+sState+"')) GROUP BY PR.REGION_CODE,ISNULL "+  " (RN.AGENCY_NAME,RN.AGENT_NAME), PR.COMPANY_CODE ORDER BY " +Orderby ;
    //If Latest ref check box is checked.
            
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
                int nTotCt =  datatab.Rows.Count ;
                dvgrid.Visible = true;
                lblcount.Visible = true;
                //
                Int32 sNullCheck = 0;
                DataView DataView1 = new DataView(datatab);
                DataRowView DataFirstRowView = null;
                DataFirstRowView = DataView1[0];

                sNullCheck = Convert.ToInt32(DataFirstRowView["SUB_COUNT"]);
                //if (string.IsNullOrEmpty(sNullCheck) || sNullCheck == "0")
                if(sNullCheck == 0)
                {
                    lblcount.Text = "No Records Found for the selected criteria !!!";
                    dvgrid.Visible = false;

                    return;
                }
                    //AddEditRows();
                lblcount.Text = "Total Record Count: " + nTotCt;
            }

            //grdHandling.DataSource = datatab;
            //grdHandling.DataBind();
            // adPolicy.Fill(dsPolicy);
            commPolicy.Connection = con;
            commPolicy.CommandText = "select PR.COMPANY_CODE AS Total_Ct,PR.COMPANY_CODE, SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE " + " (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_COUNT, SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE " + " (ST.PROD_PCNT/100)    END)   ELSE 0   END) as PAID_COUNT, SUM(CASE  WHEN (" + sNTOQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE " + "(ST.PROD_PCNT/100)    END)   ELSE 0   END) as NTO_COUNT,CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL) THEN 1 ELSE " + " (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN ISNULL " + " (PR.ANNUAL_PREMIUM,0)    ELSE ISNULL(PR.ANNUAL_PREMIUM,0)*(ST.PROD_PCNT/100)    END)   ELSE 0   END)) /(SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN " + " (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END as AVG_PAID,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE " +
" WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN " + " (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL) " + " THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END as PAID_PCNT,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN " +
" (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  " +
" PR.RATE_CLASS = 'Smoker') THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  WHEN (" + sSubQuery + ")" + " THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_NIC_PCNT,CASE WHEN (SUM(CASE  " +
" WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN " +
" (" + sPaidQuery + " AND   PR.RATE_CLASS = 'Smoker')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  " +
" WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_NIC_PCNT,SUM(CASE  WHEN " +
" (" + sSubQuery + " AND RATE_CLASS='Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as " +
" SUB_NIC_L_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE " +
" (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_L_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)  " +
" THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_L_COUNT,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'L')   THEN   (CASE  WHEN " +
" (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  PR.RATE_CLASS = 'Smoker' AND " +
" PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  WHEN (" + sSubQuery + " AND " +
" PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_L_NIC_PCNT,CASE WHEN (SUM(CASE  " +
" WHEN (" + sPaidQuery + "AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM" +
" (CASE  WHEN (" + sPaidQuery + "AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END) " +
"ELSE 0   END))/(SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   " +
" END)) end AS PAID_L_NIC_PCNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS='Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE " +
" (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NIC_G_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN " +
" (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_G_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'G')   THEN   " +
" (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_G_COUNT,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = " +
" 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  " +
" PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  " +
" WHEN (" + sSubQuery + " AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS " +
" SUB_G_NIC_PCNT, CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + " AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END) " +
" ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + " AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1" +
" ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE " +
" (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_G_NIC_PCNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS='Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN " +
" (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NIC_M_COUNT, SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS<>'Smoker' AND " +
" PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_M_COUNT,SUM(CASE  WHEN " +
" (" + sSubQuery + " AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_M_COUNT, CASE WHEN " +
" (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0" +
" ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100) " +
" END)   ELSE 0   END))  / (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   " +
" ELSE 0   END)) END AS SUB_M_NIC_PCNT,CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE " + " (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + "AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN " +
" (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/ (SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'M')   THEN   (CASE  WHEN " +
" (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_M_NIC_PCNT from POLICIES PR LEFT OUTER JOIN POLICY_SPLIT ST WITH " +
" (NOLOCK)  ON (PR.COMPANY_CODE = ST.COMPANY_CODE) and (PR.POLICY_NUMBER = ST.POLICY_NUMBER) AND  (PR.AGENT_NUMBER = ST.AGENT_NUMBER) and (PR.ISSUE_DATE = " +
" ST.SPLIT_EFF_DATE)  LEFT OUTER JOIN REGION_NAMES RN ON RN.MARKETING_COMPANY = PR.REGION_CODE where  PR.AGENT_NUMBER in (SELECT AHL0.AGENT_NUMBER FROM AGENT_HIERLIST " +
" AHL0 WHERE AHL0.COMPANY_CODE = '" + sCompany + "' AND AHL0.HIERARCHY_AGENT='" + AgentID + "') AND PR.COMPANY_CODE = '" + sCompany + "' AND  PR.RATE_CLASS IS NOT NULL AND " +
" PR.REGION_CODE <> ''  AND (('" + sState + "'='ALL')OR(PR.PI_STATE = '" + sState + "')) GROUP BY PR.COMPANY_CODE";

  
            commPolicy.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter dataadapt = new SqlDataAdapter(commPolicy.CommandText, con);
            // SqlDataAdapter dataadapter = new SqlDataAdapter(commPolicy); // Create a SQL Data Adapter and assign it the cmd value. 
            datatabTotal = new DataTable(); // Create a new Data table
            dataadapt.Fill(datatabTotal);
            //Added newly to implement Export to Excel functionality
            if (dataPolicy != null && dataPolicy.DefaultView != null)
                dataPolicy = dataPolicy.DefaultView.ToTable();
            if (datatabTotal != null && datatabTotal.Rows.Count > 0)
            {
                AddEditRows();
            }
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
                MyDataRowView["Region_Name"] = " Total Count ";
                //MyDataRowView["Region_Code"] = Convert.ToString(DataRowCurrView["Company_Code"]);
                MyDataRowView["Company_Code"] = Convert.ToString(DataRowCurrView["Company_Code"]);
                MyDataRowView["Sub_Count"] = Convert.ToString(DataRowCurrView["Sub_Count"]);
                MyDataRowView["Paid_count"] = Convert.ToString(DataRowCurrView["Paid_count"]);
                MyDataRowView["Nto_Count"] = Convert.ToString(DataRowCurrView["Nto_Count"]);
                MyDataRowView["Avg_Paid"] = Convert.ToString(DataRowCurrView["Avg_Paid"]);
                MyDataRowView["Paid_PCNT"] = Convert.ToString(DataRowCurrView["Paid_PCNT"]);
                MyDataRowView["Sub_NIC_PCNT"] = Convert.ToString(DataRowCurrView["Sub_NIC_PCNT"]);
                MyDataRowView["Paid_NIC_PCNT"] = Convert.ToString(DataRowCurrView["Paid_NIC_PCNT"]);
                MyDataRowView["Sub_NIC_L_Count"] = Convert.ToString(DataRowCurrView["Sub_NIC_L_Count"]);
                MyDataRowView["Sub_NNIC_L_Count"] = Convert.ToString(DataRowCurrView["Sub_NNIC_L_Count"]);
                MyDataRowView["Sub_L_Count"] = Convert.ToString(DataRowCurrView["Sub_L_Count"]);
                MyDataRowView["Sub_L_NIC_PCNT"] = Convert.ToString(DataRowCurrView["Sub_L_NIC_PCNT"]);
                MyDataRowView["Paid_L_NIC_PCNT"] = Convert.ToString(DataRowCurrView["Paid_L_NIC_PCNT"]);
                MyDataRowView["SUB_NIC_G_Count"] = Convert.ToString(DataRowCurrView["SUB_Nic_G_Count"]);
                MyDataRowView["Sub_NNIC_G_Count"] = Convert.ToString(DataRowCurrView["Sub_NNIC_G_Count"]);
                //
                MyDataRowView["Sub_G_Count"] = Convert.ToString(DataRowCurrView["Sub_G_Count"]);
                MyDataRowView["Sub_G_NIC_PCNT"] = Convert.ToString(DataRowCurrView["Sub_G_NIC_PCNT"]);
                MyDataRowView["Paid_G_NIC_PCNT"] = Convert.ToString(DataRowCurrView["Paid_G_NIC_PCNT"]);
                MyDataRowView["SUB_NIC_M_Count"] = Convert.ToString(DataRowCurrView["SUB_NIC_M_Count"]);
                MyDataRowView["SUB_NNIC_M_Count"] = Convert.ToString(DataRowCurrView["SUB_NNIC_M_Count"]);
                MyDataRowView["SUB_M_Count"] = Convert.ToString(DataRowCurrView["SUB_M_Count"]);
                MyDataRowView["Sub_M_NIC_PCNT"] = Convert.ToString(DataRowCurrView["Sub_M_NIC_PCNT"]);
                MyDataRowView["Paid_M_NIC_PCNT"] = Convert.ToString(DataRowCurrView["Paid_M_NIC_PCNT"]);


            }
            MyDataRowView.EndEdit();
            grInterviewsByRegion.DataSource = MyDataView1;
            //grInterviewsByRegion.DataSource = datatab;
            grInterviewsByRegion.DataBind();

        }
        protected void grInterviewsByRegion_RowDataBound(object sender, GridViewRowEventArgs e)

        {
            //string sCompany = Request.QueryString["COMPANY_CODE"];
            //string sRegion = (string)Request.QueryString["Region_Code"];
            //string sFromDate1 = (string)Request.QueryString["Fromdate"];
            //string sToDate1 = (string)Request.QueryString["Today"];
            //string sState = Request.QueryString["State"];
            string sState = "ALL";
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


            if (e.Row.RowType == DataControlRowType.DataRow)
            {

                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();

        string sRegion = e.Row.Cells[0].Text;
                //   string Policy_num = e.Row.Cells[2].Text;
                if ((e.Row.RowType == DataControlRowType.DataRow) || (e.Row.RowType == DataControlRowType.Header))
                {
                    e.Row.Cells[e.Row.Cells.Count - 1].Visible = false;
                    grInterviewsByRegion.HeaderRow.Cells[e.Row.Cells.Count - 1].Visible = false;

                }
    e.Row.Cells[1].ToolTip = "click to view details";

                string text = e.Row.Cells[0].Text;
    HyperLink link = new HyperLink();
    link.NavigateUrl = "Ratedforregion.aspx?sState=" + sState + "&Fromdate=" + FromDate + "&Today=" + ToDate +  "&COMPANY_CODE=" + sCompany + "&Region_Code=" + sRegion + "";
                link.Text = text;
                link.Target = "_blank";
                e.Row.Cells[0].Controls.Add(link);

                //PolicyView.aspx?POLICY_NUMBER=W861860006&COMPANY_CODE=16&AGENT_NUMBER=ID01001
                //e.Row.Cells[2].Text = Convert.ToString("<a href=\"PolicyView.aspx?POLICY_NUMBER="+Policy_num+"&COMPANY_CODE="+sCompany+"&AGENT_NUMBER="+Agent_num+"Target="+"_blank"+" \">"+Policy_num+"</a>");
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
            //Export("InterviewByRegion.xls", this.grInterviewsByRegion);
        }
                
        
        protected void ExportToExcel()
        {
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                dataPolicy.DefaultView.Sort = "REGION_CODE";
                dataPolicy = dataPolicy.DefaultView.ToTable();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"Rated_By_Region" + DateTime.Now.ToString();
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
