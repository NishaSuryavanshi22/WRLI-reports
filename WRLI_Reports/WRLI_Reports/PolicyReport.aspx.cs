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
using System.Configuration;

namespace WRLI_Reports
{
    public partial class _Default : System.Web.UI.Page
    {
        // SqlConnection con = new SqlConnection(CSCUtils.Utils.GetConnectionString());
        //SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
        //SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 1000");
        //   SqlConnection con = new SqlConnection("Data Source=20.15.80.160;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT;Connection Timeout = 1000");
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString);

        string agent = "WRE";
        DataSet dsPolicy = new DataSet();
        public static DataTable dataPolicy = new DataTable();
        public static string CurrselectedAgentCode = "";
        public static string CurrselectedRegionCode = "";
        public static string CurrselectedStateCode = "";
        //string CurrselectedCompanyCode = "";
        public static string CurrselectedContractCode = "";
        public string CurrselectedContractReason = "";
        public static string CurrSelectedPolStatus = "";
        string CurrDataType = "";
        //
        string FilterQuery = "";

        public static string CurrselectedCompanyCode = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                string sVal = (string)Session["Validated"];
                if (sVal == null || sVal != "A")
                {
                    //Response.Redirect("Closed.aspx");
                }
            }
            catch
            {
                Response.Redirect("Closed.aspx");
            }
            this.ddlagentNameList.SelectedIndexChanged += new EventHandler(ddlagentNameList_SelectedIndexChanged);
            this.ddlregion.SelectedIndexChanged += new EventHandler(ddlregion_SelectedIndexChanged);
            this.ddlstate.SelectedIndexChanged += new EventHandler(ddlstate_SelectedIndexChanged);
            this.ddlpolicydesc.SelectedIndexChanged += new EventHandler(ddlpolicydesc_SelectedIndexChanged);
            this.ddlpolicystatus.SelectedIndexChanged += new EventHandler(ddlpolicystatus_SelectedIndexChanged);
            this.ddlcompany.SelectedIndexChanged += new EventHandler(ddlcompany_SelectedIndexChanged);
            this.ddldatatype.SelectedIndexChanged += new EventHandler(ddldatatype_SelectedIndexChanged);
            this.ddlagent.SelectedIndexChanged += new EventHandler(ddlagent_SelectedIndexChanged);
            
            

            if (Session["LoginID"] != null)
                agent = Session["LoginID"].ToString();
            if (!IsPostBack)
            {
                tblgrid.Visible = false;
                txtTo.Text = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                txtFrom.Text = System.DateTime.Now.AddMonths(-6).ToString("d");

                string sCompany = "";
                if (Session["CompanyCode"] != null && Session["CompanyCode"].ToString() != "")
                    sCompany = Session["CompanyCode"].ToString();
                else
                    sCompany = "15";

                Session["CompanyCode"] = sCompany;
                CurrselectedCompanyCode = sCompany;
                //Commented the Connection and its added in the Utils Class
                //SqlConnection con = new SqlConnection("Data Source=BPODEV;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Initial Catalog=WRE_AGENT");
                //SqlConnection con = new SqlConnection("Data Source=WREMWEBNOR457-N;Password=Bobo$2006;Persist Security Info=True;User ID=WREAgent;Password=Bobo$2006;Initial Catalog=WRE_AGENT");
                con.Open();
                string sClause = "";
                string sTempUser = (string) Session["LocalID"];
                if ((sTempUser == "DCLERK") || (sTempUser == "JWHITFIELD"))
                {
                    sClause = "where COMPANY_CODE = 15";
                }
                SqlCommand commCompany = new SqlCommand(" select company_code,company_name from company_details "+ sClause);
                commCompany.Connection = con;
                DataSet dscomp = new DataSet();
                SqlDataAdapter adcomp = new SqlDataAdapter(commCompany.CommandText, con);
                adcomp.Fill(dscomp);
                List<string> lstcomp = new List<string>();
                int iCompanySelected = 0;
                for (int i = 0; i < dscomp.Tables[0].Rows.Count; i++)
                {
                    string sCompanyNumber = dscomp.Tables[0].Rows[i].ItemArray[0].ToString();
                    string sCompanyName = dscomp.Tables[0].Rows[i].ItemArray[1].ToString();
                    lstcomp.Add(sCompanyNumber + " - " + sCompanyName);
                    if (sCompanyNumber == sCompany)
                    {
                        iCompanySelected = i;
                    }


                }
                //ddlcompany.DataSource = lstcomp;
                //ddlcompany.DataBind();
                //
                if (dscomp != null)
                    dscomp.Tables[0].DefaultView.Sort = "company_code ASC";
                //ddlcompany.DataSource = dscomp;
                //ddlcompany.DataTextField = "company_name";
                //ddlcompany.DataValueField = "company_code";
                ddlcompany.DataSource = lstcomp;
                ddlcompany.DataBind();
                //ddlcompany.SelectedIndex = 2;
                //

                ddlcompany.SelectedIndex = iCompanySelected;

                string sCompanySelect = "";
                string sSelectedCompany = "";
                if (ddlcompany.SelectedValue != null)
                {
                    string[] NameValuePair = ddlcompany.SelectedItem.Value.ToString().Split('-');
                    //string[] NameValuePair = ddlcompany.SelectedItem.Text.Split('-');
                    if (NameValuePair != null)
                    {
                        sCompanySelect = " COMPANY_CODE = '" + NameValuePair[0].Trim().ToString() + "' AND ";
                        sSelectedCompany = NameValuePair[0].Trim().ToString();
                    }

                }
                //Agent
                SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,ISNULL(DISPLAYNAME,'') AS DISPLAYNAME,SORTNAME FROM AGENT_HIERLIST WHERE "+sCompanySelect+" HIERARCHY_AGENT = '" + agent + "' ORDER BY SORTNAME ASC");
                comm.Connection = con;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(comm.CommandText, con);
                ad.Fill(ds);
                if (ds!=null)
                    ds.Tables[0].DefaultView.Sort = "SORTNAME ASC";

                List<string> lstagent = new List<string>();
                List<string> lstagentName = new List<string>();

                
                //ddlagentNameList.Items.Add("ALL");

                ddlagentNameList.Items.Clear();
                //ddlagentNameList.Items.Add("ALL");
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    string sName = ds.Tables[0].Rows[i].ItemArray[1].ToString().Trim();
                    if (sName != "")
                    {
                        //lstagent.Add(ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() + "-" + sName + " (" + ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() + ")");
                        lstagent.Add ( sName + " (" + ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() + ")");
                        //Adding the Agent name to the dropdown
                        ddlagentNameList.Items.Add(sName + "-" + ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() );
                    }
                }
                //Bind the agent name and code to the list.
                ddlagent.DataSource = lstagent;
                ddlagent.DataBind();
                if (ddlagent.Items.FindByValue("ALL") == null)
                {
                    ddlagent.Items.Insert(0, new ListItem("ALL", "ALL"));
                    if (CurrselectedAgentCode == "")
                    {
                        ddlagent.SelectedIndex = 0;
                    }

                }
                //State
                SqlCommand commStates = new SqlCommand("SELECT * FROM ALLSTATES");
                commStates.Connection = con;
                DataSet dsStates = new DataSet();
                SqlDataAdapter adStates = new SqlDataAdapter(commStates.CommandText, con);
                adStates.Fill(dsStates);
                if (dsStates != null) 
                    dsStates.Tables[0].DefaultView.Sort = "STATE_ABBR ASC";
                ddlstate.Items.Clear();
                ddlstate.DataSource = dsStates;
                ddlstate.DataTextField = "STATE_NAME";
                ddlstate.DataValueField = "STATE_ABBR";
                ddlstate.DataBind();

                ddlstate.Items.Insert(0, "ALL");
                ddlstate.SelectedIndex = 0;

                if (ddlstate.Items.FindByValue("ALL") == null)
                {
                    ddlstate.Items.Insert(0, new ListItem("ALL", "ALL"));
                    if (CurrselectedStateCode == "")
                    {
                        ddlstate.SelectedIndex = 0;
                    }

                }

                SqlCommand commReason = new SqlCommand("SELECT DISTINCT CONTRACT_REASON, UPPER(CONTRACT_DESC) AS CONTRACT_DESC FROM POLICIES2 WHERE RTRIM(ISNULL(CONTRACT_REASON,''))<>'' ORDER BY UPPER(CONTRACT_DESC)");
                commReason.Connection = con;
                DataSet dsReason = new DataSet();
                SqlDataAdapter adReason = new SqlDataAdapter(commReason.CommandText, con);
                //ALL
                adReason.Fill(dsReason);
                ddlpolicydesc.Items.Clear();
                ddlpolicydesc.DataSource = dsReason;
                ddlpolicydesc.DataTextField = "CONTRACT_DESC";
                ddlpolicydesc.DataValueField = "CONTRACT_REASON";
                ddlpolicydesc.DataBind();
                if (ddlpolicydesc.Items.FindByValue("ALL") == null)
                {
                    ddlpolicydesc.Items.Insert(0, new ListItem("ALL", "ALL"));
                    ddlpolicydesc.SelectedIndex = 0;
                }


                SqlCommand commMarket = new SqlCommand("SELECT ISNULL(MARKETING_COMPANY,'UNKNOWN') AS MARKETING_COMPANY,* FROM REGION_NAMES RN "+
                                                       "WHERE (RN.MARKETING_COMPANY IN (SELECT REGION_CODE FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = '"+agent+"' and company_code = '"+sSelectedCompany+"'))");
                commMarket.Connection = con;
                DataSet dsMarket = new DataSet();
                SqlDataAdapter adMarket = new SqlDataAdapter(commMarket.CommandText, con);
                adMarket.Fill(dsMarket);
                if (dsMarket != null)
                    dsMarket.Tables[0].DefaultView.Sort = "MARKETING_COMPANY ASC";
                ddlregion.Items.Clear();
                ddlregion.DataSource = dsMarket;
                ddlregion.DataTextField = "MARKETING_COMPANY";
                //Newly added siva
                ddlregion.DataValueField = "MARKETING_COMPANY";
                ddlregion.DataBind();
                if (ddlregion.Items.FindByValue("ALL") == null)
                {
                    ddlregion.Items.Insert(0, new ListItem("ALL", "ALL"));
                    if (CurrselectedRegionCode == "")
                    {
                        ddlregion.SelectedIndex = 0;
                    }
                }
                //ddlregion.Items[0].Selected = true;
                con.Close();
                GetDropdownVales();
            }
        }
        //Issue Fixing
        void ddldatatype_SelectedIndexChanged(object sender, EventArgs e)
        {
                CurrDataType = ddldatatype.SelectedValue.ToString();
            //throw new NotImplementedException();
        }

        void ddlcompany_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] NameValuePair = new string[2];
            if (!ddlcompany.SelectedItem.ToString().ToUpper().Equals("ALL"))
            //CurrselectedCompanyCode = ddlcompany.SelectedValue.ToString();
            {
                tblgrid.Visible = false;
                if (ddlcompany.SelectedValue != null)
                {
                    //NameValuePair = ddlcompany.SelectedItem.Value.ToString().Split['-'];
                    NameValuePair = ddlcompany.SelectedItem.Text.Split('-');
                    if (NameValuePair != null)
                    {
                        CurrselectedCompanyCode = NameValuePair[0].Trim().ToString();
                        SqlCommand comm = new SqlCommand("SELECT AGENT_NUMBER,ISNULL(DISPLAYNAME,'') AS DISPLAYNAME,SORTNAME FROM AGENT_HIERLIST WHERE COMPANY_CODE = " + CurrselectedCompanyCode + " AND HIERARCHY_AGENT = '" + agent + "' ORDER BY SORTNAME ASC");
                        comm.Connection = con;
                        DataSet ds = new DataSet();
                        SqlDataAdapter ad = new SqlDataAdapter(comm.CommandText, con);
                        ad.Fill(ds);
                        if (ds != null)
                            ds.Tables[0].DefaultView.Sort = "SORTNAME ASC";

                        List<string> lstagent = new List<string>();
                        List<string> lstagentName = new List<string>();
                        ddlagentNameList.Items.Clear();
                        ddlagentNameList.Items.Add("ALL");
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            string sName = ds.Tables[0].Rows[i].ItemArray[1].ToString().Trim();
                            if (sName != "")
                            {
                                lstagent.Add(ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() + "-" + sName);
                                //Adding the Agent name to the dropdown

                                ddlagentNameList.Items.Add(sName + " (" + ds.Tables[0].Rows[i].ItemArray[0].ToString().Trim() + ")");
                            }
                        }
                        //Bind the agent name and code to the list.
                        ddlagent.DataSource = lstagent;
                        ddlagent.DataBind();
                        SqlCommand commMarket = new SqlCommand("SELECT ISNULL(MARKETING_COMPANY,'UNKNOWN') AS MARKETING_COMPANY,* FROM REGION_NAMES RN " +
                                                               "WHERE (RN.MARKETING_COMPANY IN (SELECT REGION_CODE FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = '" + agent + "' and company_code = '" + CurrselectedCompanyCode + "'))");
                        commMarket.Connection = con;
                        DataSet dsMarket = new DataSet();
                        SqlDataAdapter adMarket = new SqlDataAdapter(commMarket.CommandText, con);
                        adMarket.Fill(dsMarket);
                        if (dsMarket != null)
                            dsMarket.Tables[0].DefaultView.Sort = "MARKETING_COMPANY ASC";
                        ddlregion.Items.Clear();
                        ddlregion.DataSource = dsMarket;
                        ddlregion.DataTextField = "MARKETING_COMPANY";
                        //Newly added siva
                        ddlregion.DataValueField = "MARKETING_COMPANY";
                        ddlregion.DataBind();
                        if (ddlregion.Items.FindByValue("ALL") == null)
                        {
                            ddlregion.Items.Insert(0, new ListItem("ALL", "ALL"));
                            if (CurrselectedRegionCode == "")
                            {
                                ddlregion.SelectedIndex = 0;
                            }
                        }
                    }

                }
            }
            //throw new NotImplementedException();
        }

        void ddlpolicystatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            CurrSelectedPolStatus = "";
            if (! ddlpolicystatus.SelectedItem.ToString().ToUpper().Equals("ALL"))
                CurrSelectedPolStatus = ddlpolicystatus.SelectedValue.ToString();
            //throw new NotImplementedException();
        }

        void ddlpolicydesc_SelectedIndexChanged(object sender, EventArgs e)
        {
            CurrselectedContractCode = "";
            if(! ddlpolicydesc.SelectedItem.ToString().ToUpper().Equals("ALL"))
                CurrselectedContractCode = ddlpolicydesc.SelectedValue.ToString();
               //throw new NotImplementedException();
        }

        void ddlstate_SelectedIndexChanged(object sender, EventArgs e)
        {
            CurrselectedStateCode = "";
            if (! ddlstate.SelectedItem.ToString().ToUpper().Equals("ALL"))
                CurrselectedStateCode = ddlstate.SelectedValue.ToString();
        }

        void ddlregion_SelectedIndexChanged(object sender, EventArgs e)
        {
            CurrselectedRegionCode = "";
            if (! ddlregion.SelectedItem.ToString().ToUpper().Equals("ALL"))
                CurrselectedRegionCode = ddlregion.SelectedValue.ToString();
            //throw new NotImplementedException();
        }
        void ddlagent_SelectedIndexChanged(object sender, EventArgs e)
        {
            int nAgentCodect = 0;
            int nAgentNameList = 0;
            //string sAgentName = "";
            string sCurrSelectedItem = "";
            string[] AgentNamearray = new string[2];
            string[] AgentCodeArray = new string[2];
            if (ddlagent != null)
                nAgentCodect = ddlagent.Items.Count;
            if (nAgentCodect > 0)
            {
                if (ddlagent.SelectedItem.ToString().Trim().ToUpper().Equals("ALL"))
                {
                    CurrselectedAgentCode = "";
                    return;
                }
                
                sCurrSelectedItem = ddlagent.SelectedItem.Text;
                AgentCodeArray = sCurrSelectedItem.Split('(');
                if (ddlagentNameList != null && ddlagentNameList.Items.Count > 0)
                {
                    nAgentNameList = ddlagentNameList.Items.Count;
                    for (int nIndex = 0; nIndex < nAgentNameList; nIndex++)
                    {
                        AgentNamearray = ddlagentNameList.Items[nIndex].ToString().Split('-');

                        if (AgentNamearray != null && AgentNamearray.Length == 2 && AgentCodeArray != null && AgentCodeArray[0].Trim().Equals(AgentNamearray[0].ToString().Trim()))
                        {
                            if (AgentNamearray[1] != null)
                                CurrselectedAgentCode = AgentNamearray[1].ToString().Trim();

                        }
                    }
                }

            }
        }

        //Issue Fixing
        void ddlagentNameList_SelectedIndexChanged(object sender, EventArgs e)
        {
            /* int nAgentCodect = 0;
            string sAgentName ="";
            string [] AgentNamearray = new string[2];
            if(ddlagent != null)
                nAgentCodect = ddlagent.Items.Count;
            if (ddlagentNameList != null && ddlagentNameList.Items.Count > 0)
            {
                for (int nIndex = 0; nIndex < nAgentCodect; nIndex++)
                {
                    AgentNamearray = ddlagent.Items[nIndex].Text.Split('-');
                    if (AgentNamearray != null && AgentNamearray.Length ==2 )
                    {
                        if (ddlagentNameList.SelectedItem.Text.Trim().Equals(AgentNamearray[1].ToString().Trim()))
                        {
                            if (AgentNamearray[0] != null)
                            {
                                CurrselectedAgentCode = AgentNamearray[0].ToString().Trim();
                                break;
                            }
                        }
                    }
                }

            } */
            //throw new NotImplementedException();
        }


        protected void InvokeSP()
        {
            string selectedComp = "ALL";
            string selectedAgent = "ALL";
            string frmDate = Utils.DateToLPDate(txtFrom.Text);

            string tDate = Utils.DateToLPDate(txtTo.Text);

            con.Open();

            int index = ddlcompany.SelectedItem.Value.LastIndexOf("-");
            if (index > 0)
            {
                selectedComp = ddlcompany.SelectedItem.Value.Substring(0, index);
            }

            int indexagent = ddlagent.SelectedItem.Value.LastIndexOf("-");
            if (indexagent > 0)
            {
                selectedAgent = ddlagent.SelectedItem.Value.Substring(0, indexagent);
            }
            string sSql = "select dbo.GET_POLICY_INSURED(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS INSURED_NAME," +
                "RIGHT(RTRIM(ISNULL(PI_SOC_SEC_NUMBER,'0000')),4) AS INSURED_SSN," +
                "dbo.GET_POLICY_OWNER(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS OWNER_NAME," +
                "dbo.GET_POLICY_PAYOR(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS PAYOR_NAME," +
                "PO.COMPANY_CODE," +
                "PO.POLICY_NUMBER," +
                "dbo.[GET_BASE_PRODUCT_DESCEX](ISNULL(PO.PRODUCT_CODE,''),ISNULL(PP.PRODUCT_CODE,'')) AS PLAN_TYPE," +
                "po.RATE_CLASS as TOBACCO_USE," +
                "dbo.GET_POLICY_ISSUE_STATE(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS ISSUE_STATE," +
                "dbo.GET_POLICY_PI_CITY(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS INSURED_CITY," +
                "ISNULL(PI_STATE,dbo.GET_POLICY_PI_STATE(PO.COMPANY_CODE,PO.POLICY_NUMBER)) AS INSURED_STATE," +
                "dbo.LPDATE_TO_STRDATE(ISNULL(dbo.GET_POLICY_PI_DOB(PO.COMPANY_CODE,PO.POLICY_NUMBER),'')) AS INSURED_DOB," +
                "dbo.GET_POLICY_ISSUE_AGE(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS ISSUE_AGE," +
                "dbo.FORMAT_PHONE(ISNULL(dbo.GET_POLICY_PI_PHONE(PO.COMPANY_CODE,PO.POLICY_NUMBER),'0000000000')) AS INSURED_PHONE," +
                "ISNULL(po.FACE_AMOUNT,0) AS FACE_AMOUNT," +
                "ISNULL(pp.MODE_PREMIUM,'0') AS MODE_PREMIUM," +
                "UPPER(dbo.GETTRANSLATION('POLICY INFO:BILLING_MODE',RTRIM(CASE WHEN (ISNULL(BILLING_MODE,'')<10) THEN '0'+RTRIM(ISNULL(BILLING_MODE,'')) ELSE RTRIM(ISNULL(BILLING_MODE,'')) END))) " +
                "AS BILLING_MODE," +
                "dbo.GET_POLICY_BILLING_FORM(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS BILLING_FORM," +
                "PO.CONTRACT_CODE AS STATUS," +
                "CASE WHEN (PO.CONTRACT_CODE = 'A') THEN 'ACTIVE' ELSE PO.CONTRACT_DESC END AS STATUS_DESC," +
                "dbo.GET_AGENT_DISPLAY_NAME(PO.COMPANY_CODE,PO.AGENT_NUMBER,'L') AS SERVICE_AGENT_NAME," +
                "PO.AGENT_NUMBER AS SERVICE_AGENT," +
                "dbo.GET_AGENT_STATUS(PO.COMPANY_CODE,PO.AGENT_NUMBER) AS SA_STATUS," +
                "CASE WHEN (ISNULL(PO.REGION_CODE,'0')='0') THEN dbo.GETAGENTREGION(PO.AGENT_NUMBER,PO.COMPANY_CODE) ELSE PO.REGION_CODE END " +
                "AS SA_REGION_CODE," +
                "dbo.GET_AGENT_CO_PHONE(PO.COMPANY_CODE,PO.AGENT_NUMBER) AS SERVICE_AGENT_PHONE," +
                "dbo.GET_AGENT_DISPLAY_NAME(PO.COMPANY_CODE,ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER),'L') AS WRITING_AGENT_NAME," +
                "ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER) AS WRITING_AGENT," +
                "dbo.GET_AGENT_STATUS(PO.COMPANY_CODE,ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER)) AS WA_STATUS," +
                "dbo.GETAGENTREGION(ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER),PO.COMPANY_CODE) AS WA_REGION_CODE," +
                "dbo.GET_AGENT_CO_PHONE(PO.COMPANY_CODE,ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER)) AS WRITING_AGENT_PHONE," +
                "PO.DURATION AS DURATION," +
                "dbo.LPDATE_TO_STRDATE(ISNULL(PO.ISSUE_DATE,'')) AS ISSUE_DATE," +
                "dbo.LPDATE_TO_STRDATE(ISNULL(PO.PAID_TO_DATE,'')) AS PAID_TO_DATE, " +
                "0 AS CASH_VALUE," +
                "dbo.LPDATE_TO_STRDATE(application_date) as APPLICATION_SIGNED_DATE," +
                "dbo.LPDATE_TO_STRDATE(po.app_received_date) as APPLICATION_RECEIVED_DATE," +
                "case when (po.contract_code='T') then dbo.LPDATE_TO_STRDATE(po.last_change_date) else dbo.LPDATE_TO_STRDATE('') end as TERMINATION_DATE," +
                "dbo.GET_POLICY_PI_GENDER(PO.COMPANY_CODE,PO.POLICY_NUMBER) AS Insured_gender " +
                "from POLICIES2 PO left outer join pending_policy pp on po.policy_number = pp.policy_number and po.company_code=pp.company_code WHERE (PO.AGENT_NUMBER IN (SELECT AGENT_NUMBER FROM AGENT_HIERLIST WHERE HIERARCHY_AGENT = '" + agent + "' and company_code = '" + CurrselectedCompanyCode.ToString() + "' OR '" + CurrselectedCompanyCode.ToString() + "'='ALL')) AND (po.APP_RECEIVED_DATE BETWEEN '" + frmDate + "' and '" + tDate + "' ) AND " + FilterQuery; // +"ORDER BY POLICY_NUMBER DESC";
                
                
                
                //(PO.COMPANY_CODE = '" + selectedComp + "' OR '" + selectedComp + "'='ALL') AND (PO.AGENT_NUMBER = '" + selectedAgent + "' OR '" + ddlregion.SelectedItem.Value + "'='ALL') AND (PI_STATE = '" + ddlstate.SelectedItem.Value + "' OR '" + ddlregion.SelectedItem.Value + "' = 'ALL')AND (po.CONTRACT_CODE = '" + ddlpolicystatus.SelectedItem.Value + "' OR '" + ddlregion.SelectedItem.Value + "' = 'ALL') AND (CONTRACT_REASON = '" + ddlpolicydesc.SelectedItem.Value + "' OR '" + ddlregion.SelectedItem.Value + "' = 'ALL') AND (PO.REGION_CODE = '" + ddlregion.SelectedItem.Value + "' OR '" + ddlregion.SelectedItem.Value + "' = 'ALL') ";//ORDER BY POLICY_NUMBER DESC";
            Label1.Text = sSql;
            Label1.Visible = false ;
            SqlCommand commPolicy = new SqlCommand(sSql);
            commPolicy.Connection = con;

            SqlDataAdapter adPolicy = new SqlDataAdapter(commPolicy.CommandText, con);
            adPolicy.Fill(dsPolicy);
            con.Close();

            if (dsPolicy.Tables[0].Rows.Count == 0)
            {
                DataTable dt = new DataTable();
                DataColumn dc = new DataColumn();
                #region columncount
                if (dt.Columns.Count == 0)
                {
                    dt.Columns.Add("INSURED NAME", typeof(string));
                    dt.Columns.Add("INSURED SSN", typeof(string));
                    dt.Columns.Add("OWNER NAME", typeof(string));
                    dt.Columns.Add("PAYER NAME", typeof(string));
                    dt.Columns.Add("COMPANY CODE", typeof(string));
                    dt.Columns.Add("POLICY NUMBER", typeof(string));
                    dt.Columns.Add("PLAN", typeof(string));
                    dt.Columns.Add("TOBACCO USE", typeof(string));
                    dt.Columns.Add("ISSUE STATE", typeof(string));
                    dt.Columns.Add("INSURED CITY", typeof(string));
                    dt.Columns.Add("INSURED STATE", typeof(string));
                    dt.Columns.Add("INSURED DOB", typeof(string));
                    dt.Columns.Add("ISSUE AGE", typeof(string));
                    dt.Columns.Add("INSURED_PHONE", typeof(string));
                    dt.Columns.Add("FACE AMT/MO.INC", typeof(string));
                    dt.Columns.Add("MODE PREMIUM", typeof(string));
                    dt.Columns.Add("BILLING MODE", typeof(string));
                    dt.Columns.Add("BILLING FORM", typeof(string));
                    dt.Columns.Add("STATUS", typeof(string));
                    dt.Columns.Add("STATUS DESC", typeof(string));
                    dt.Columns.Add("SERVICE AGENT NAME", typeof(string));
                    dt.Columns.Add("SERVICE AGENT", typeof(string));
                    dt.Columns.Add("SERVICE STATUS", typeof(string));
                    dt.Columns.Add("SERVICE REGION", typeof(string));
                    dt.Columns.Add("SERVICE AGENT PHONE", typeof(string));
                    dt.Columns.Add("WRITING AGENT NAME", typeof(string));
                    dt.Columns.Add("WRITING AGENT", typeof(string));
                    dt.Columns.Add("WRITING STATUS", typeof(string));
                    dt.Columns.Add("WRITING REGION", typeof(string));
                    dt.Columns.Add("WRITING AGENT PHONE", typeof(string));
                    dt.Columns.Add("DURATION", typeof(string));
                    dt.Columns.Add("ISSUE DATE", typeof(string));
                    dt.Columns.Add("PAID TO DATE", typeof(string));
                    dt.Columns.Add("CASH VALUE", typeof(string));
                    dt.Columns.Add("APPLICATION SIGNED DATE", typeof(string));
                    dt.Columns.Add("APPLICATION RECEIVED DATE", typeof(string));
                    dt.Columns.Add("TERMINATION DATE", typeof(string));
                    dt.Columns.Add("INSURED GENDER", typeof(string));
                }
                #endregion
                dvgrid.Style.Add("height", "120px");
                DataRow NewRow = dt.NewRow();
                dt.Rows.Add(NewRow);
                DataView dv = new DataView(dt);
                try
                {
                    dv.Sort = "POLICY_NUMBER";
                }
                catch { }
                GridView1.DataSource = dv;

                GridView1.DataBind();
                lblcount.Text = "No Records Found for the selected criteria !!";

            }
            else
            {
                dvgrid.Style.Add("height", "600px");
                DataView dv = new DataView(dsPolicy.Tables[0]);
                dv.Sort = "POLICY_NUMBER";
                GridView1.DataSource = dv;
                GridView1.DataBind();
                if (dsPolicy.Tables[0] != null)
                {
                    lblcount.Text = dsPolicy.Tables[0].Rows.Count.ToString();
                    dataPolicy = dsPolicy.Tables[0];
                    //Session.Add("dataPolicy", dataPolicy);
                }
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            tblgrid.Visible = true;
            GetDropdownVales();
            GetSelectedValues();
            InvokeSP();
            
        }

        //Issue Fixing
        protected void GetDropdownVales()
        {
            CurrselectedContractCode = "";
            CurrselectedStateCode = "";
            CurrselectedRegionCode = "";
            //CurrselectedAgentCode = "";
            if (! ddlpolicydesc.SelectedItem.ToString().ToUpper().Equals("ALL"))
                CurrselectedContractCode = ddlpolicydesc.SelectedValue.ToString().Trim();
            //
            if ( ddlstate.SelectedItem != null && ! ddlstate.SelectedItem.ToString().ToUpper().Equals("ALL"))
                CurrselectedStateCode = ddlstate.SelectedValue.ToString().Trim();
            //
            if (!ddlregion.SelectedItem.ToString().ToUpper().Equals("ALL"))
                CurrselectedRegionCode = ddlregion.SelectedValue.ToString().Trim();
            //

           /* int nAgentCodect = 0;
            string sAgentName = "";
            string[] AgentNamearray = new string[2];
            if (ddlagent != null)
                nAgentCodect = ddlagent.Items.Count;
            if (nAgentCodect > 0)
            {
                AgentNamearray = ddlagent.SelectedItem.Text.Split('-');
                if (AgentNamearray != null && AgentNamearray.Length == 2)
                {
                    if (AgentNamearray[0] != null)
                        CurrselectedAgentCode = AgentNamearray[0].ToString().Trim();

                }

            }  */
            
        }
        //Issue Fixing
        protected void GetSelectedValues()
        {
            
            if(CurrselectedCompanyCode != null && CurrselectedCompanyCode.Trim().Length > 0)
                FilterQuery = " PO.COMPANY_CODE = '" + CurrselectedCompanyCode.ToString() + "'" + " AND ";

            if (CurrselectedAgentCode != null && CurrselectedAgentCode.Trim().Length > 0)
                FilterQuery += "((PO.AGENT_NUMBER = '" + CurrselectedAgentCode + "') or " +
                    "(ISNULL(dbo.GET_POLICY_WRITING_AGENT(PO.COMPANY_CODE,PO.POLICY_NUMBER),PO.AGENT_NUMBER)= '"+CurrselectedAgentCode+"')) AND ";

            if (CurrselectedStateCode != null && CurrselectedStateCode.Trim().Length > 0)
                FilterQuery += "PO.PI_STATE  = '" + CurrselectedStateCode.ToString() + "'" + " AND ";

            if (CurrselectedContractCode != null && CurrselectedContractCode.Trim().Length > 0)
                FilterQuery += "PO.CONTRACT_CODE  = '" + CurrselectedContractCode.ToString() + "'" + " AND ";

            if (CurrselectedRegionCode != null && CurrselectedRegionCode.Trim().Length > 0)
                FilterQuery += "PO.REGION_CODE  = '" + CurrselectedRegionCode.ToString() + "'";

            /* if (CurrDataType != null && CurrDataType.Trim().Length > 0)
                FilterQuery += "PO.DATA_TYPE  = '" + CurrDataType + "'"; */
            

                        //string source = "My name is Marco and I'm from Italy";
            string[] stringSeparators = new string[] {"AND"};
            var result = FilterQuery.Split(stringSeparators, StringSplitOptions.None);
            string LastAnd = "";
            string OriginalFilterQuery = "";
            if (result != null && result.Length > 0)
            {
                int nSplitCount = result.Length - 1;
                LastAnd = result[nSplitCount - 1].ToString();
                //if (LastAnd.Trim().Equals("AND"))
                //{
                for (int nIndex = 0; nIndex <= nSplitCount; nIndex++)
                    {
                        if (nIndex == 0)
                            OriginalFilterQuery += result[nIndex].Trim().ToString();
                        else if (nIndex == nSplitCount)
                        {
                            if (result[nIndex].Trim().ToString().Length > 0)
                                OriginalFilterQuery += " AND " + result[nIndex].Trim().ToString();
                        }
                        else
                            OriginalFilterQuery += " AND " + result[nIndex].Trim().ToString();
                        
                    }
                    FilterQuery = OriginalFilterQuery;

               // }
                
            }
                

        }

        protected void ddlagent_PreRender(object sender, EventArgs e)
        {
            /*
             * if (ddlagent.Items.FindByValue(string.Empty) == null)
            {
                ddlagent.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
             * */
        }
        

        protected void ddlregion_PreRender(object sender, EventArgs e)
        {
            /*
            if (ddlregion.Items.FindByValue(string.Empty) == null)
            {
                ddlregion.Items.Insert(0, new ListItem("ALL", "ALL"));
                if (CurrselectedRegionCode == "")
                {
                    ddlregion.SelectedIndex = 0;
                }
            }
             * */
        }


        protected void ddlstate_PreRender(object sender, EventArgs e)
        {
            /*
            if (ddlstate.Items.FindByValue(string.Empty) == null)
            {
                ddlstate.Items.Insert(0, new ListItem("ALL", "ALL"));
                if (CurrselectedStateCode == "")
                {
                    ddlstate.SelectedIndex = 0;
                }

            }
             * */
        }


        protected void ddlpolicydesc_PreRender(object sender, EventArgs e)
        {
            /*
            if (ddlpolicydesc.Items.FindByValue(string.Empty) == null)
            {
                ddlpolicydesc.Items.Insert(0, new ListItem("ALL", "ALL"));
            }
             * */
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            //Export("PolicyReport.xls", this.GridView1);
            ExportToExcel();


        }
        protected void ExportToExcel()
        {
            //InvokeSP();
//            dataPolicy = Session["dataPolicy"] as DataTable;
            dataPolicy.DefaultView.Sort = "POLICY_NUMBER";
            dataPolicy = dataPolicy.DefaultView.ToTable();
            //dataPolicy = GridView1.DataSource as DataTable;
            if (dataPolicy.Rows.Count > 0 && dataPolicy != null)
            {
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string filename = @"Policy_Report_" + DateTime.Now.ToString("yyyyMMdd");
                Response.AddHeader("Content-Disposition", "inline;filename=" + filename.Replace("/", "").Replace(":", "") + ".xlsx");
                //Call  Export function
                Response.BinaryWrite(Utils.ExportToCSVFileOpenXML(dataPolicy));

                Response.Flush();
                Response.End();
            }

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

        protected void GridView1_SelectedIndexChanging(object sender, GridViewSelectEventArgs e)
        {

        }

        protected void GridView1_PageIndexChanged(object sender, EventArgs e)
        {

        }

        protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {

        }
       
    }

}
