using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;

using System.Text.RegularExpressions;
using System.Configuration;
namespace WRLI_Reports
{
    //PolicyView.aspx?POLICY_NUMBER=W861860006&COMPANY_CODE=16&AGENT_NUMBER=ID01001
    public partial class PolicyView : System.Web.UI.Page
    {
        string sCustomer = "";
        string sANum = "";
        string sAName = "";
        public string FixDate(string sDate)
        {
            string sResult = "";
            if ((sDate == "1/1/1900") || (sDate == "0") || (sDate == "") || (sDate == null))
            {
                sResult = "";
            }
            else
            {

                sResult = sDate.Substring(4, 2) + "/" +
                          sDate.Substring(6, 2) + "/" +
                          sDate.Substring(0, 4);
            }
            return sResult;
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            string sCompany = Request.QueryString["COMPANY_CODE"];
            string sPolicyNum = (string) Request.QueryString["POLICY_NUMBER"];
            string sAgentNumber = (string)Request.QueryString["AGENT_NUMBER"];
            string sSql = "SELECT AH.TELE_NUM as AGENT_PHONE, PO.PAID_TO_DATE, PO.ISSUE_STATE, PO.PRODUCT_CODE,PO.PRODUCT_DESC, PO.FACE_AMOUNT, ISSUE_AGE, " +
                        "PO.MODE_PREMIUM, PO.ANNUAL_PREMIUM, PO.BILLING_MODE, PO.BILLING_FORM, " +
                        "RTRIM(LTRIM(PO.OWNER_LAST)) + ', ' + RTRIM(LTRIM(PO.OWNER_FIRST)) as OWNER_NAME, " +
                        "RTRIM(LTRIM(PO.BEN_LAST)) + ', ' + RTRIM(LTRIM(PO.BEN_FIRST)) as BEN_NAME, " +
                        "RTRIM(LTRIM(PO.PAY_LAST)) + ', ' + RTRIM(LTRIM(PO.PAY_FIRST)) as PAY_NAME, " +
                        "RTRIM(LTRIM(PO.SA_LAST)) + ', ' + RTRIM(LTRIM(PO.SA_FIRST)) + ' ('+RTRIM(LTRIM(SERVICE_AGENT))+')' as SA_NAME, " +
                        "PO.DELIVERY_DATE,RTRIM(LTRIM(PO.PI_LAST)) + ', ' + RTRIM(LTRIM(PO.PI_FIRST)) as INS_NAME, " +
                        "APPLICATION_DATE,PO.CONTRACT_CODE,  dbo.LPDATE_TO_STRDATE(PO.PI_DOB) as ULP_PI_DOB, PO.ISSUE_DATE,PP.RECEIVED_DATETIME, LINE_OF_BUSINESS, PO.PRODUCT_CODE, " +
                        "dbo.FORMAT_PHONE(ISNULL(PO.PI_PHONE, '')) AS PI_PHONEX, PO.*,PO.CONTRACT_DESC AS CONTRACT_DESC_EX,  " +
                        "PO.CONTRACT_DESC AS CONTRACT_REASON, PO.DELIVERY_DATE, PO.RATE_CLASS,AGENCY_NAME, SA_REGION_CODE, " +
                        "PO.OWNER_ADDRESS1, PO.OWNER_ADDRESS2, PO.OWNER_CITY, PO.OWNER_STATE, PO.OWNER_ZIP, PO.OWNER_PHONE, " +
                        "ISNULL(PP.INSURED_ADDRESS_TYPE, '0') AS XINSURED_ADDRESS_TYPE, ISNULL(PP.INSURED_ADDRESS_LINE1, '') AS INSURED_ADDRESS_LINE1, " +
                        "ISNULL(PP.INSURED_ADDRESS_LINE1, '') AS INSURED_ADDRESS_LINE2, ISNULL(PP.INSURED_ADDRESS_CITY, '') AS INSURED_ADDRESS_CITY, " +
                        "ISNULL(PP.INSURED_ADDRESS_STATE, '') AS INSURED_ADDRESS_STATE, ISNULL(PP.INSURED_ADDRESS_ZIP, '') AS INSURED_ADDRESS_ZIP, " +
                        "dbo.FORMAT_PHONE(ISNULL(PP.INSURED_TELEPHONE, '')) AS XINSURED_TELEPHONE, ISNULL(PP.EFT_ROUTING_NUMBER, '') AS EFT_ROUTING_NUMBER, " +
                        "ISNULL(PP.EFT_ACCOUNT_NUMBER, '') AS EFT_ACCOUNT_NUMBER, ISNULL(PP.EFT_ACCOUNT_TYPE, '') AS EFT_ACCOUNT_TYPE, " +
                        "ISNULL(PP.EFT_DRAFT_START_DATE, '') AS EFT_DRAFT_START_DATE, " +
                        "ISNULL(PP.RELATIONSHIP_OF_BENEFICIARY, '') AS RELATIONSHIP_OF_BENEFICIARY, " +
                        "ISNULL(PP.APPROVING_UNDERWRITER, '') AS APPROVING_UNDERWRITER, ISNULL(PP.REISSUE_INDICATOR, '') AS REISSUE_INDICATOR, " +
                        "ISNULL(PP.RIDER, '') AS RIDER, ISNULL(PP.REPLACEMENT_INDICATOR, '') AS REPLACEMENT_INDICATOR, " +
                        "ISNULL(PP.REPLACEMENT_COMPANY_NAME, '') AS REPLACEMENT_COMPANY_NAME, " +
                        "dbo.GETTRANSLATION('POLICY INFO:PAID_UP_TYPE', PAID_UP_TYPE) AS PUPT, PRODUCT_DESC " +
                        "FROM POLICY AS PO " +
                        "LEFT OUTER JOIN POLICY_SPLIT PS " +
                        "on PS.POLICY_NUMBER = PO.POLICY_NUMBER and PS.COMPANY_CODE = PO.COMPANY_CODE " +
                        "LEFT OUTER JOIN REGION_NAMES RN " +
                        "ON SA_REGION_CODE = MARKETING_COMPANY " +
                        "LEFT OUTER JOIN PENDING_POLICY PP " +
                        "ON PP.COMPANY_CODE = PO.COMPANY_CODE AND PP.POLICY_NUMBER = PO.POLICY_NUMBER " +
                        "LEFT OUTER JOIN AGENTS AH " +
                        "ON AH.AGENT_NUMBER = PP.AGENT_NUMBER AND AH.COMPANY_CODE = PP.COMPANY_CODE " +
                        "WHERE PO.POLICY_NUMBER = '[POLICY_NUMBER]' AND " +
                        "(WRITING_AGENT = '[WRITING_AGENT]' OR " +
                        "SERVICE_AGENT = '[SERVICE_AGENT]' OR "+
                        "PS.AGENT_NUMBER = '[AGENT_NUMBER]')"; 
            string sSql1 = "SELECT AH.TELE_NUM as AGENT_PHONE, '' as PAID_TO_DATE,'' AS PAY_NAME, ISNULL(PP.INSURED_ADDRESS_STATE, '') AS ISSUE_STATE, '' as PRODUCT_DESC,  " +
                "PP.PRODUCT_CODE,PP.FACE_AMOUNT,'' AS BEN_NAME, ISSUE_AGE, PP.STATUS_CODE AS CONTRACT_CODE, WFWORKSTEPNAME as CONTRACT_DESC, " +
                "PP.MODE_PREMIUM, PP.ANNUAL_PREMIUM, '' as BILLING_MODE, '' as BILLING_FORM, " +
                " ISNULL(AH.AGENT_NUMBER, PP.AGENT_NUMBER) AS AGENT_NUMBER, dbo.FORMAT_PHONE(ISNULL(TELE_NUM_OFFICE, '')) AS SA_PHONE, " +
                "AH.STATUS_CODE as SA_STATUS, AH.NAME_FORMAT_CODE as SA_FORMAT, AH.NAME_BUSINESS as SA_BUSINESS,  " +
                "PP.AGENT_NUMBER as SERVICE_AGENT, RTRIM(LTRIM(AH.INDIVIDUAL_FIRST)) + ' ' + RTRIM(LTRIM(AH.INDIVIDUAL_LAST)) + ' ('+RTRIM(LTRIM(AH.AGENT_NUMBER))+')' as SA_NAME,   " +
                "RTRIM(LTRIM(PP.INDIVIDUAL_LAST)) + ', ' + RTRIM(LTRIM(PP.INDIVIDUAL_FIRST)) AS INS_NAME, '' AS OWNER_NAME, "+
                "PP.STATUS_CODE AS PENDING_STATUS, AGENCY_NAME, AH.REGION_CODE as SA_REGION_CODE, " +
                "'' as ULP_PI_DOB,'' as OWNER_ADDRESS1, '' as OWNER_ADDRESS2, '' as OWNER_CITY, '' as OWNER_STATE, '' AS OWNER_ZIP, '' AS OWNER_PHONE, " +
                "CASE WHEN PP.DELIVERY_DATE IS NOT NULL THEN 'Policy Mailed' ELSE PP.WFWORKSTEPNAME END AS PENDING_DESC, " +
                "ISNULL(PP.INSURED_ADDRESS_TYPE, '0') AS XINSURED_ADDRESS_TYPE, ISNULL(PP.INSURED_ADDRESS_LINE1, '') AS INSURED_ADDRESS_LINE1, " +
                "ISNULL(PP.INSURED_ADDRESS_LINE1, '') AS INSURED_ADDRESS_LINE2, ISNULL(PP.INSURED_ADDRESS_CITY, '') AS INSURED_ADDRESS_CITY, " +
                "ISNULL(PP.INSURED_ADDRESS_STATE, '') AS INSURED_ADDRESS_STATE, ISNULL(PP.INSURED_ADDRESS_ZIP, '') AS INSURED_ADDRESS_ZIP, " +
                "dbo.FORMAT_PHONE(ISNULL(PP.INSURED_TELEPHONE, '')) AS XINSURED_TELEPHONE, ISNULL(PP.EFT_ROUTING_NUMBER, '') AS EFT_ROUTING_NUMBER, " +
                "ISNULL(PP.EFT_ACCOUNT_NUMBER, '') AS EFT_ACCOUNT_NUMBER, ISNULL(PP.EFT_ACCOUNT_TYPE, '') AS EFT_ACCOUNT_TYPE, " +
                "ISNULL(PP.EFT_DRAFT_START_DATE, '') AS EFT_DRAFT_START_DATE, " +
                "ISNULL(PP.RELATIONSHIP_OF_BENEFICIARY, '') AS RELATIONSHIP_OF_BENEFICIARY, " +
                "ISNULL(PP.APPROVING_UNDERWRITER, '') AS APPROVING_UNDERWRITER, ISNULL(PP.REISSUE_INDICATOR, '') AS REISSUE_INDICATOR, " +
                "ISNULL(PP.RIDER, '') AS RIDER, ISNULL(PP.REPLACEMENT_INDICATOR, '') AS REPLACEMENT_INDICATOR, " +
                "ISNULL(PP.REPLACEMENT_COMPANY_NAME, '') AS REPLACEMENT_COMPANY_NAME, ISNULL(PP.RATE_CLASS, '') AS RATE_CLASS, " +
                "claims.issue_date, APP_RECEIVED_DATE as APPLICATION_DATE " +
                "FROM PENDING_POLICY PP " +
                "left outer join claims_reporting claims " +
                "on pp.policy_number = claims.policy_number and pp.company_code = claims.company_code " +
                "LEFT OUTER JOIN AGENTS AH " +
                "ON AH.AGENT_NUMBER = PP.AGENT_NUMBER AND AH.COMPANY_CODE = PP.COMPANY_CODE " +
                "LEFT OUTER JOIN REGION_NAMES RN " +
                "ON AH.REGION_CODE = MARKETING_COMPANY " +
                "WHERE PP.POLICY_NUMBER = '[POLICY_NUMBER]' AND " +
                "(PP.AGENT_NUMBER = '[AGENT_NUMBER]')";
            bool bLoadData = false;
            //    string sDBConnectionString = CSCUtils.Utils.GetConnectionString();
            // SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 3600");
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

            //   using (SqlConnection conn = new SqlConnection(sDBConnectionString))
            {
                try
                {
                    sSql = sSql.Replace("[POLICY_NUMBER]",sPolicyNum);
                    sSql = sSql.Replace("[AGENT_NUMBER]", sAgentNumber);
                    sSql = sSql.Replace("[WRITING_AGENT]", sAgentNumber);
                    sSql = sSql.Replace("[SERVICE_AGENT]", sAgentNumber);
                    sSql1 = sSql1.Replace("[POLICY_NUMBER]", sPolicyNum);
                    sSql1 = sSql1.Replace("[AGENT_NUMBER]", sAgentNumber);
                    sSql1 = sSql1.Replace("[WRITING_AGENT]", sAgentNumber);
                    sSql1 = sSql1.Replace("[SERVICE_AGENT]", sAgentNumber);
                    con.Open();
                    SqlCommand cm = new SqlCommand(sSql, con);
                    SqlDataReader rsPolicyInfo = cm.ExecuteReader();

                    if (rsPolicyInfo.HasRows)
                    {
                        rsPolicyInfo.Read();
                        bLoadData = true;
                    }
                    else
                    {
                        con.Close();
                        con.Dispose();
                        con.Open();
                        SqlCommand cm1 = new SqlCommand(sSql1, con);
                        rsPolicyInfo = cm1.ExecuteReader();
                        if (rsPolicyInfo.HasRows)
                        {
                            rsPolicyInfo.Read();
                            bLoadData = true;
                        }

                    }
                    if (!bLoadData)
                    {
                        pnlData.Visible = false;
                    }
                    else
                    {
                        pnlData.Visible = true;
                        string sInsured = rsPolicyInfo["INS_NAME"].ToString().Trim();
                        sCustomer = sInsured;
                        string sOwner = rsPolicyInfo["OWNER_NAME"].ToString().Trim();
                        lblName.Text = sInsured;
                        sPolicyNum = rsPolicyInfo["POLICY_NUMBER"].ToString().Trim();
                        lblPolicyNumber.Text = sPolicyNum;
                        string sContractCode = rsPolicyInfo["CONTRACT_CODE"].ToString().Trim();
                        string sEffDate = "";
                        if ((sContractCode) == "P")
                        {
                            sEffDate = "Not Mentioned";
                        } else
                        {
                            sEffDate = FixDate(rsPolicyInfo["ISSUE_DATE"].ToString());
                        }
                        string sAppDate = (rsPolicyInfo["APPLICATION_DATE"].ToString().Trim());
                        sAppDate = FixDate(sAppDate);
                        string sAppRECDate = FixDate(rsPolicyInfo["APP_RECEIVED_DATE"].ToString());
                        string sTimeStamp = (rsPolicyInfo["RECEIVED_DATETIME"].ToString().Trim());
                        string sIssueDate = FixDate(rsPolicyInfo["APP_RECEIVED_DATE"].ToString());
                        string sMailedDate = FixDate(rsPolicyInfo["DELIVERY_DATE"].ToString());
                        string sContractDesc = rsPolicyInfo["CONTRACT_DESC"].ToString().Trim();
                        lblAppDate.Text = sAppDate;
                        lblAppRecDate.Text = sAppRECDate;
                        lblAppTS.Text = sTimeStamp;
                        lblIssueDate.Text = sEffDate;
                        lblMailedDate.Text = sMailedDate;
                        lblStatus.Text = sContractCode;
                        lblStatusDesc.Text = sContractDesc;
                        if (sContractDesc == "")
                        {
                            lblStatusDesc.Visible = false;
                            lblStatus1.Visible = false;
                        }
                        if (sContractCode=="T")
                        {
                            lblStatusDesc.CssClass = "redtexttxt";
                        }
                        string sRate = rsPolicyInfo["RATE_CLASS"].ToString();
                        if (sRate != "")
                        {
                            if (sRate == "Smoker")
                            {
                                sRate = "Yes";
                            }
                            else
                            {
                                sRate = "No";
                            }
                        }
                        lblTabacco.Text = sRate;
                        lblIName.Text = sInsured;
                        lblOName.Text = sOwner;
                        string sAddress1 = rsPolicyInfo["INSURED_ADDRESS_LINE1"].ToString().Trim();
                        string sAddress2 = rsPolicyInfo["INSURED_ADDRESS_CITY"].ToString().Trim() + ", " +
                                           rsPolicyInfo["INSURED_ADDRESS_STATE"].ToString().Trim() + "  " +
                                           rsPolicyInfo["INSURED_ADDRESS_ZIP"].ToString().Trim();
                        lblIAddress1.Text = sAddress1;
                        lblIAddress2.Text = sAddress2;
                        string sOAddress1 = rsPolicyInfo["OWNER_ADDRESS1"].ToString().Trim();
                        string sOAddress2 = rsPolicyInfo["OWNER_CITY"].ToString().Trim() + ", " +
                                           rsPolicyInfo["OWNER_STATE"].ToString().Trim() + "  " +
                                           rsPolicyInfo["OWNER_ZIP"].ToString().Trim();
                        lblOAddress1.Text = sOAddress1;
                        lblOAddress2.Text = sOAddress2;
                        string sIPhone = rsPolicyInfo["XINSURED_TELEPHONE"].ToString().Trim();
                        string sOPhone = rsPolicyInfo["OWNER_PHONE"].ToString().Trim();
                        lblIPhone.Text = sIPhone;
                        lblOPhone.Text = sOPhone;
                        string sDOB = (rsPolicyInfo["ULP_PI_DOB"].ToString().Trim());
                        string sAge = rsPolicyInfo["ISSUE_AGE"].ToString().Trim();
                        lblDOB.Text = sDOB;
                        if (sAge == "")
                        {
                            DateTime dtAge = DateTime.Parse(sDOB);
                            int iAge = DateTime.Today.Year - dtAge.Year;
                            if (dtAge > DateTime.Today.AddYears(-iAge))
                                iAge--;
                            sAge = iAge.ToString();
                        }
                        lblAge.Text = sAge;
                        string sBenny = rsPolicyInfo["BEN_NAME"].ToString().Trim();
                        if (sBenny == ",")
                        {
                            sBenny = sInsured;
                        }
                        lblBenny.Text = sBenny;
                        string sFace = rsPolicyInfo["FACE_AMOUNT"].ToString().Trim();
                        if (sFace != "")
                        {
                            sFace = "$" + sFace.Replace("$", "");
                        }
                        lblFace.Text = sFace;
                        string sProductCode = rsPolicyInfo["PRODUCT_CODE"].ToString().Trim();
                        string sPlanCode = (rsPolicyInfo["PRODUCT_DESC"]).ToString().Trim();
                        if (sPlanCode == "")
                        {
                            string sTemp = (rsPolicyInfo["PLAN_CODE"]).ToString().Trim();
                            if (sTemp == "L")
                            {
                                sPlanCode = "Level Benefit";
                            }
                            if (sTemp == "M")
                            {
                                sPlanCode = "Modified Benefit";
                            }

                            if (sTemp == "G")
                            {
                                sPlanCode = "Graded Benefit";

                            }
                        }
                        lblBasePlan.Text = sProductCode;
                        lblPlan.Text = sPlanCode;
                        string sIssueState = (rsPolicyInfo["ISSUE_STATE"]).ToString().Trim();
                        if (sIssueState=="")
                            sIssueState = (rsPolicyInfo["INSURED_ADDRESS_STATE"]).ToString().Trim();
                        lblIssueState.Text = sIssueState;
                        string sPayName = (rsPolicyInfo["PAY_NAME"]).ToString().Trim();
                        lblPayor.Text = sPayName;
                        string sModePremium = (rsPolicyInfo["MODE_PREMIUM"]).ToString().Trim();
                        string sAnnualPremium = (rsPolicyInfo["ANNUAL_PREMIUM"]).ToString().Trim();
                        string sPaymentMode = (rsPolicyInfo["BILLING_MODE"]).ToString().Trim();
                        string sPaymentForm = (rsPolicyInfo["BILLING_FORM"]).ToString().Trim();
                        if (sModePremium != "")
                            sModePremium = "$" + sModePremium;
                        if (sAnnualPremium != "")
                            sAnnualPremium = "$" + sAnnualPremium;
                        switch (sPaymentForm)
                        {
                            case "0":
                                sPaymentForm = "Direct";
                                break;
                            case "H":
                                sPaymentForm = "List Bill";
                                break;
                            case "G":
                                sPaymentForm = "Pre Authorized Check";
                                break;
                            default:
                                sPaymentForm = "Direct";
                                break;
                        }
                        switch (sPaymentMode)
                        {
                            case "1":
                                sPaymentMode = "Monthly";
                                break;
                            case "3":
                                sPaymentMode = "Quarterly";
                                break;
                            case "6":
                                sPaymentMode = "Semi - Annual";
                                break;
                            default:
                                sPaymentMode = "Annual";
                                break;
                        }
                        lblModalPremium.Text = sModePremium;
                        lblAnnualPremium.Text = sAnnualPremium;
                        lblPaymentMode.Text = sPaymentMode;
                        lblPaymentForm.Text = sPaymentForm;
                        string sPaidToDate = FixDate((rsPolicyInfo["PAID_TO_DATE"]).ToString().Trim());
                        lblPaidToDate.Text = sPaidToDate;
                        string sServiceAgent = (rsPolicyInfo["SA_NAME"]).ToString().Trim();
                        lblServiceAgent.Text = sServiceAgent;
                        string sAgency = (rsPolicyInfo["AGENCY_NAME"]).ToString().Trim();
                        string sRegion = (rsPolicyInfo["SA_REGION_CODE"]).ToString().Trim();
                        sANum = sAgentNumber;
                        sAName = sServiceAgent;
                        lblServiceAgency.Text = sAgency + "(" + sRegion + ")";
                        string sAgentPhone = (rsPolicyInfo["AGENT_PHONE"]).ToString().Trim();
                        //sAgentPhone = String.Format("{0:(###) ###-####}", sAgentPhone);
                        sAgentPhone = Regex.Replace(sAgentPhone, @"(\d{3})(\d{3})(\d{4})", "($1) $2-$3");
                        lblSAPhone.Text = sAgentPhone;

                        //gridview
                        con.Close();
                        con.Open();
                        String sql_grid = " SELECT distinct INDIVIDUAL_FIRST + ' ' +  INDIVIDUAL_LAST as OWNER_NAME,TRANSLATION.DESCRIPTION AS REQ_DESCRIPTION,RECEIPT_FLAG AS UND_FLAG, COMMENT AS COMMENTS,dbo.LPDATE_TO_STRDATE(ADD_DATE) AS UND_DATE,ADD_DATE,MET_DATE AS UND_O_DATE, RECEIPT_FLAG FROM PENDING_REQUIREMENTS PR LEFT OUTER JOIN TRANSLATION on TRANS_NAME = 'PENDING NB:REQUIREMENT TYPE' AND CODE = RECORD_TYPE WHERE PR.COMPANY_CODE = '" + sCompany + "' AND PR.POLICY_NUMBER = '" + sPolicyNum + "'  ORDER BY ADD_DATE ASC";
                        SqlDataAdapter da = new SqlDataAdapter(sql_grid, con);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        GridView1.DataSource = dt;
                        GridView1.DataBind();

                        con.Close();



                    }
                }
                catch (Exception ex) { }
            }
        }

        protected void btnEMSI_Click(object sender, EventArgs e)
        {
            Response.Write("<script>");
            Response.Write("window.open('https://emsionline.emsinet.com/','_blank')");
            Response.Write("</script>");
        }

        protected void btnSource_Click(object sender, EventArgs e)
        {
            string sPolicyNum = (string)Request.QueryString["POLICY_NUMBER"];
            string sURL = "https://sams.1sourceaccess.com/autoLogin.aspx?username=clico&password=clico123&policy=" + sPolicyNum;
            Response.Write("<script>");
            Response.Write("window.open('"+sURL+"','_blank')");
            Response.Write("</script>");
        }

        protected void btnEmail_Click(object sender, EventArgs e)
        {
            string sAgentEmail = (string)Session["emailAddr"];
            string sPolicyNum = (string)Request.QueryString["POLICY_NUMBER"];
            var sTemp = "Pending - " + sPolicyNum + " - " + sCustomer + " (" + sANum + " - " + sAName + ")";
            string sURL = "Email.aspx?RETURN=close&to=nb-uw@csc.com&from=" + sAgentEmail + "&subject=" + sTemp;
            Response.Write("<script>");
            Response.Write("window.open('" + sURL + "','_blank')");
            Response.Write("</script>");

        }
    }
}