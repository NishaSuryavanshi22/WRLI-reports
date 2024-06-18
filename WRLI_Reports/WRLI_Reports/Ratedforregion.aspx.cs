//using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Configuration;
using DocumentFormat.OpenXml.VariantTypes;
using CSCUtils;
using System.IO;

namespace WRLI_Reports
{
    public partial class Ratedforregion : System.Web.UI.Page
    {
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);


        DataSet dsPolicy = new DataSet();
        public static DataTable dataPolicy = new DataTable();
        string sCompany = "15";
        string sRegionCode = "INS";
        string AgentID = "WRE";
        string FromDate = string.Empty;
        string ToDate = string.Empty;
        string Orderby = "Agent_number ASC";
        string OrderDir = "ASC";
        bool bNet = false;
        bool bRegion = false;
        int nRowct = 0;
        string sType = "type";
        Int32[] arr_NB = new Int32[] { };
        string[] strRowValue = new string[3];
        DataTable datatab = new DataTable(); // Create a new Data table
        DataTable datatabTotal = new DataTable();

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

            string sPaidColumn = "PAYMENT_DATE";
            string sCompany = Request.QueryString["COMPANY_CODE"];
            string sRegion = (string)Request.QueryString["Region_Code"];
            string sFromDate = (string)Request.QueryString["Fromdate"];
            string sToDate = (string)Request.QueryString["Today"];
            string sState = (string)Request.QueryString["sState"];

            string sSubQuery = "((((PR." + sPaidColumn + " BETWEEN " + sFromDate + " AND " + sToDate + ")AND(PR.RECORD_TYPE='I'))OR((PR.APP_RECEIVED_DATE BETWEEN " + sFromDate + " AND " + sToDate + ")AND(PR.RECORD_TYPE='P'))) AND (PR.APP_RECEIVED_DATE IS NOT NULL))";
            string sPaidQuery = "((PR." + sPaidColumn + " BETWEEN " + sFromDate + " AND " + sToDate + ") AND   (PR.PAYMENT_FLAG = 'Y') AND ((PR.CONTRACT_CODE = 'A') OR (PR.CONTRACT_CODE = 'T')OR (PR.CONTRACT_CODE = 'S')))";
            string sNTOQuery = "((PR." + sPaidColumn + " BETWEEN " + sFromDate + " AND " + sToDate + ") AND   (PR.PAYMENT_FLAG = 'Y') AND ((PR.CONTRACT_CODE = 'T')AND(PR.CONTRACT_REASON = 'NT')))";



            SqlCommand commPolicy = new SqlCommand();

            commPolicy.Connection = con;
            //commPolicy.CommandType = CommandType.StoredProcedure;

            commPolicy.CommandText = "select  PR.AGENT_NUMBER,dbo.GET_AGENT_DISPLAY_NAME(PR.COMPANY_CODE,PR.AGENT_NUMBER,'L') AS AGENT_NAME,PR.COMPANY_CODE, SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_COUNT, SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as PAID_COUNT, SUM(CASE  WHEN (" + sNTOQuery  + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as NTO_COUNT,CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN ISNULL(PR.ANNUAL_PREMIUM,0)    ELSE ISNULL(PR.ANNUAL_PREMIUM,0)*(ST.PROD_PCNT/100)    END)   ELSE 0   END)) /(SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END as AVG_PAID,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END as PAID_PCNT,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  PR.RATE_CLASS = 'Smoker')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_NIC_PCNT,CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + " AND   PR.RATE_CLASS = 'Smoker')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_NIC_PCNT,SUM(CASE  WHEN (" + sSubQuery +" AND RATE_CLASS='Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NIC_L_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_L_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_L_COUNT,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_L_NIC_PCNT,CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + "AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_L_NIC_PCNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS='Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NIC_G_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_G_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_G_COUNT,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_G_NIC_PCNT, CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + " AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + " AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_G_NIC_PCNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS='Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NIC_M_COUNT, SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_M_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_M_COUNT, CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  / (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_M_NIC_PCNT,CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + "AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/ (SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_M_NIC_PCNT from POLICIES PR LEFT OUTER JOIN POLICY_SPLIT ST WITH (NOLOCK)  ON (PR.COMPANY_CODE = ST.COMPANY_CODE) and (PR.POLICY_NUMBER = ST.POLICY_NUMBER) AND  (PR.AGENT_NUMBER = ST.AGENT_NUMBER) and (PR.ISSUE_DATE = ST.ISSUE_DATE)  LEFT OUTER JOIN REGION_NAMES RN ON RN.MARKETING_COMPANY = PR.REGION_CODE " +
                "where  PR.AGENT_NUMBER in (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE COMPANY_CODE = '" + sCompany + "' AND HIERARCHY_AGENT='" + AgentID + "') AND PR.COMPANY_CODE = '" + sCompany + "' AND  PR.RATE_CLASS IS NOT NULL AND PR.REGION_CODE <> '' AND PR.REGION_CODE = '" + sRegionCode + "' AND (('" + sState + "'='ALL')OR(PR.PI_STATE = '" + sState + "')) GROUP BY PR.AGENT_NUMBER,dbo.GET_AGENT_DISPLAY_NAME(PR.COMPANY_CODE,PR.AGENT_NUMBER,'L'), PR.COMPANY_CODE ORDER BY "  +Orderby+"";

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
                int nTotCt = datatab.Rows.Count;
                dvgrid.Visible = true;
                lblcount.Visible = true;
                //
                Int32 sNullCheck = 0;
                DataView DataView1 = new DataView(datatab);
                DataRowView DataFirstRowView = null;
                DataFirstRowView = DataView1[0];

               // sNullCheck = Convert.ToInt32(DataFirstRowView["SUB_COUNT"]);
                //if (string.IsNullOrEmpty(sNullCheck) || sNullCheck == "0")
                //if (sNullCheck == 0)
                //{
                 //   lblcount.Text = "No Records Found for the selected criteria !!!";
                 //   dvgrid.Visible = false;

                //    return;
              //  }
                //AddEditRows();
                lblcount.Text = "Total Record Count: " + nTotCt;
                regioncode.Text = sRegion;

            }
            commPolicy.Connection = con;
            commPolicy.CommandText =  "select PR.COMPANY_CODE, SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_COUNT, SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as PAID_COUNT, SUM(CASE  WHEN (" + sNTOQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as NTO_COUNT,CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN ISNULL(PR.ANNUAL_PREMIUM,0)    ELSE ISNULL(PR.ANNUAL_PREMIUM,0)*(ST.PROD_PCNT/100)    END)   ELSE 0   END)) /(SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END as AVG_PAID,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END as PAID_PCNT,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  PR.RATE_CLASS = 'Smoker')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  WHEN (" + sSubQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_NIC_PCNT,CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + " AND   PR.RATE_CLASS = 'Smoker')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN (" + sPaidQuery + ")   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_NIC_PCNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS='Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NIC_L_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_L_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_L_COUNT,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_L_NIC_PCNT,CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + "AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'L')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_L_NIC_PCNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS='Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NIC_G_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_G_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_G_COUNT,CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  /(SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_G_NIC_PCNT, CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + " AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + " AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/(SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'G')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_G_NIC_PCNT,SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS='Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NIC_M_COUNT, SUM(CASE  WHEN (" + sSubQuery + " AND RATE_CLASS<>'Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_NNIC_M_COUNT,SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END) as SUB_M_COUNT, CASE WHEN (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 THEN 0 ELSE (SUM(CASE  WHEN (" + sSubQuery + " AND  PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))  / (SUM(CASE  WHEN (" + sSubQuery + " AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) END AS SUB_M_NIC_PCNT,CASE WHEN (SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) = 0 then 0 ELSE (SUM(CASE  WHEN (" + sPaidQuery + "AND   PR.RATE_CLASS = 'Smoker' AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END))/ (SUM(CASE  WHEN (" + sPaidQuery + "AND PLAN_CODE = 'M')   THEN   (CASE  WHEN (ST.PROD_PCNT IS NULL)    THEN 1    ELSE (ST.PROD_PCNT/100)    END)   ELSE 0   END)) end AS PAID_M_NIC_PCNT from POLICIES PR LEFT OUTER JOIN POLICY_SPLIT ST WITH (NOLOCK)  ON (PR.COMPANY_CODE = ST.COMPANY_CODE) and (PR.POLICY_NUMBER = ST.POLICY_NUMBER) AND  (PR.AGENT_NUMBER = ST.AGENT_NUMBER) and (PR.ISSUE_DATE = ST.SPLIT_EFF_DATE)  LEFT OUTER JOIN REGION_NAMES RN ON RN.MARKETING_COMPANY = PR.REGION_CODE where  PR.AGENT_NUMBER in (SELECT AHL0.AGENT_NUMBER FROM AGENT_HIERLIST AHL0 WHERE AHL0.COMPANY_CODE = '" + sCompany + "' AND AHL0.HIERARCHY_AGENT='" + AgentID + "') AND PR.COMPANY_CODE = '" + sCompany + "' AND  PR.RATE_CLASS IS NOT NULL AND PR.REGION_CODE = '" + sRegionCode + "' AND PR.REGION_CODE <> '' AND (('" + sState + "'='ALL')OR(PR.PI_STATE = '" + sState + "')) GROUP BY PR.COMPANY_CODE";

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


            //grdHandling.DataSource = datatab;
            //grdHandling.DataBind();

            // adPolicy.Fill(dsPolicy);

            con.Close();

        }
        protected void InitGridColumns(int Rowcount)
        {
            arr_NB[0] = new Int32();
            arr_NB[1] = new Int32();
            arr_NB[2] = new Int32();
            arr_NB[3] = new Int32();
            //arr_NB[4] = new Int32();

        }

        protected void AddEditRows()
        {
            DataView MyDataView1 = new DataView(datatab);
            DataView MyDataView2 = new DataView(datatabTotal);
            DataRowView DataRowCurrView = null;
            MyDataView1.AllowNew = true;

            //MyDataRowView["active"] = 111;
            //MyDataRowView["sub"] = 222;

            nRowct = datatab.Rows.Count;
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

            for (int nIndex = 0; nIndex < 1; nIndex++)
            {
                DataRowCurrView = MyDataView2[nIndex];

                //ReadColumnValues(DataRowCurrView, ref arr_NB);
                // End of new logic

                position = i + 1; //Dont want to insert at the row, but after.
                                  //if (FilterResultsType == "1")
                MyDataRowView["Agent_Name"] = " Total Count ";
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
            grdratedforregion.DataSource = MyDataView1;
            //grInterviewsByRegion.DataSource = datatab;
            grdratedforregion.DataBind();

        }

        protected void grdratedforregion_RowDataBound(object sender, GridViewRowEventArgs e)

        {
            string sCompany = Request.QueryString["COMPANY_CODE"];
            string sRegion = (string)Request.QueryString["Region_Code"];
            string sFromDate1 = (string)Request.QueryString["Fromdate"];
            string sToDate1 = (string)Request.QueryString["Today"];
            string sState ="ALL";
            string StausQuery = "&STATE=" + sState;
         
            if (e.Row.RowType == DataControlRowType.DataRow)
            {

                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();

                string Agent_number = e.Row.Cells[0].Text;
                //   string Policy_num = e.Row.Cells[2].Text;
                if ((e.Row.RowType == DataControlRowType.DataRow) || (e.Row.RowType == DataControlRowType.Header))
                {
                    e.Row.Cells[e.Row.Cells.Count - 1].Visible = false;
                    grdratedforregion.HeaderRow.Cells[e.Row.Cells.Count - 1].Visible = false;

                }
                e.Row.Cells[0].ToolTip = "click to view details";

                string text = e.Row.Cells[0].Text;
                HyperLink link = new HyperLink();
                link.NavigateUrl = "RatedforAgent.aspx?sState=" + sState + "&Fromdate=" + sFromDate1 + "&Today=" + sToDate1 + "&COMPANY_CODE=" + sCompany + "&Agentid=" + Agent_number + "&np=1=" + StausQuery + "";
                link.Text = text;
                link.Target = "_blank";
                e.Row.Cells[0].Controls.Add(link);

                //PolicyView.aspx?POLICY_NUMBER=W861860006&COMPANY_CODE=16&AGENT_NUMBER=ID01001
                //e.Row.Cells[2].Text = Convert.ToString("<a href=\"PolicyView.aspx?POLICY_NUMBER="+Policy_num+"&COMPANY_CODE="+sCompany+"&AGENT_NUMBER="+Agent_num+"Target="+"_blank"+" \">"+Policy_num+"</a>");
            }

        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            string sPaidColumn = "PAYMENT_DATE";
            string sCompany = Request.QueryString["COMPANY_CODE"];
            string sRegion = (string)Request.QueryString["Region_Code"];
            string sFromDate = (string)Request.QueryString["Fromdate"];
            string sToDate = (string)Request.QueryString["Today"];
            string sState = (string)Request.QueryString["sState"];
            Response.Redirect("RatedByRegion.aspx?Region_Code=" + sRegion + "&Fromdate=" + sFromDate + "&Today=" + sToDate + "&sState=" + sState + "&COMPANY_CODE=" + sCompany + "");

        }
        protected void Button2_Click(object sender, EventArgs e)
        {
            //Export the grid data to excel without hitting the Database
            ExportToExcel();
            //Export("ClaimsByRegion.xls", this.grRequirementsByRegion);
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