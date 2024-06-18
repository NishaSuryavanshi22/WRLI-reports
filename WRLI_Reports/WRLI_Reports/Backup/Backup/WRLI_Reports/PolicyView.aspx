<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PolicyView.aspx.cs" Inherits="WRLI_Reports.PolicyView" %>
<%@ Register TagPrefix="uc1" TagName="customMenu" Src="~/StyleStuff.ascx" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Policy View</title>
        <link rel="stylesheet" runat="server" media="screen" href="wrli1.css" />
    <style type="text/css">
        .auto-style1 {
            height: 44px;
        }
        .auto-style7 {
            height: 217px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
          <uc1:customMenu ID="style1" runat="server" />
          <asp:Panel runat="server" ID="pnlData">
            <table width="100%" class="auto-style1" >
                <tr style="width:100%">
                    <td align="left" 
                        style="font-family: arial, Helvetica, sans-serif; color: white" 
                        bgcolor="#0099FF" 
                        class="auto-style4"><b>Policy Information</b></td>
                </tr> 
                <tr >
                    <td class="auto-style7">

                        <asp:Label ID="lblName" runat="server" CssClass="bluetexthdr" Text="LastName, FirstName"></asp:Label>
                        <br />

                        <asp:Label ID="Label1" runat="server" Text="Policy Number:" CssClass="bluetexttxt"></asp:Label>

                        &nbsp;<asp:Label ID="lblPolicyNumber" runat="server" CssClass="bluetexttxt" Text="W123456789"></asp:Label>
                        <br />
                        <asp:Label ID="Label3" runat="server" CssClass="bluetexttxt" Text="Application Date:"></asp:Label>
                        &nbsp;<asp:Label ID="lblAppDate" runat="server" CssClass="bluetexttxt" Text="W123456789"></asp:Label>
                        <br />
                        <asp:Label ID="Label5" runat="server" CssClass="bluetexttxt" Text="Application Received Date:"></asp:Label>
                        &nbsp;<asp:Label ID="lblAppRecDate" runat="server" CssClass="bluetexttxt" Text="12/07/2018 "></asp:Label>
                        <br />
                        <asp:Label ID="Label7" runat="server" CssClass="bluetexttxt" Text="Application Timestamp:"></asp:Label>
                        &nbsp;<asp:Label ID="lblAppTS" runat="server" CssClass="bluetexttxt" Text="12/07/2018 2:24 PM "></asp:Label>
                        <br />
                        <asp:Label ID="Label9" runat="server" CssClass="bluetexttxt" Text="Issue Date:"></asp:Label>
                        &nbsp;<asp:Label ID="lblIssueDate" runat="server" CssClass="bluetexttxt" Text="01/02/2019"></asp:Label>
                        <br />
                        <asp:Label ID="Label11" runat="server" CssClass="bluetexttxt" Text="Mailed Date:"></asp:Label>
                        &nbsp;<asp:Label ID="lblMailedDate" runat="server" CssClass="bluetexttxt" Text="12/11/2018"></asp:Label>
                        <br />
                        <asp:Label ID="Label13" runat="server" CssClass="bluetexttxt" Text="Status Code:"></asp:Label>
                        &nbsp;<asp:Label ID="lblStatus" runat="server" CssClass="bluetexttxt" Text="T"></asp:Label>
                        &nbsp;<asp:Label ID="lblStatus1" runat="server" CssClass="bluetexttxt" Text="-"></asp:Label>
                        &nbsp;<asp:Label ID="lblStatusDesc" runat="server" CssClass="bluetexttxt" Text="NOT TAKEN"></asp:Label>
                        <br />
                        <asp:Label ID="Label15" runat="server" CssClass="bluetexttxt" Text="Tabacco Use:"></asp:Label>
                        &nbsp;<asp:Label ID="lblTabacco" runat="server" CssClass="bluetexttxt" Text="No"></asp:Label>

                    </td>
                </tr>
             </table>
             <table style="border-style:Solid;border-color:Gray;border-width: 1px;border-frame: box" BORDER="0" CELLPADDING="1" CELLSPACING="0" width="100%">
	            <tr bgcolor='#6699CC'>
		            <td colspan="4" align="center"><font color="white"><strong>General Information</strong></font></td>
	            </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label16" runat="server" Text="Insured Name: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblIName" runat="server" CssClass="bluetexttxt" Text="LastName, FirstName"></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label18" runat="server" Text="Owner Name: " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblOName" runat="server" CssClass="bluetexttxt" Text="LastName, FirstName"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right" valign="top">
                        <asp:Label ID="Label6" runat="server" Text="Address: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left" valign="top">
                        &nbsp;<asp:Label ID="lblIAddress1" runat="server" CssClass="bluetexttxt" Text="Address"></asp:Label><br />
                        &nbsp;<asp:Label ID="lblIAddress2" runat="server" CssClass="bluetexttxt" Text="Address2"></asp:Label>
                    </td>
                    <td align="right" valign="top">
                        <asp:Label ID="Label10" runat="server" Text="Address: " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left" valign="top">
                        &nbsp;<asp:Label ID="lblOAddress1" runat="server" CssClass="bluetexttxt" Text="Address"></asp:Label><br />
                        &nbsp;<asp:Label ID="lblOAddress2" runat="server" CssClass="bluetexttxt" Text="Address2"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label8" runat="server" Text="Phone: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblIPhone" runat="server" CssClass="bluetexttxt" Text="(405) 555-1212"></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label14" runat="server" Text="Phone: " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblOPhone" runat="server" CssClass="bluetexttxt" Text=" "></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label12" runat="server" Text="Date of Birth: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblDOB" runat="server" Text="10/25/1979" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label19" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label23" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label21" runat="server" Text="Current Age: " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblAge" runat="server" Text="39" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label17" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label20" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label24" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label25" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label26" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label27" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label2" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label4" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label22" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label54" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label32" runat="server" Text="Beneficiary: " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblBenny" runat="server" Text="LastName,FirstName" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label34" runat="server" Text="Base Plan Code: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblBasePlan" runat="server" Text="015S1C" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label36" runat="server" Text="Face Amount: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblFace" runat="server" Text="$100,000.00" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label33" runat="server" Text="Plan: " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblPlan" runat="server" Text="10 YR Term CCR" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label56" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label58" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label61" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label62" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label28" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label29" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label30" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label31" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label35" runat="server" Text="Issue State: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblIssueState" runat="server" Text="VA" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label38" runat="server" Text="Payor: " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblPayor" runat="server" Text="LastName, FirstName" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label37" runat="server" Text="Modal Premium: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblModalPremium" runat="server" Text="$100.80" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label44" runat="server" Text="Payment Mode: " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblPaymentMode" runat="server" Text="Monthly" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label39" runat="server" Text="Annualized Premium: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblAnnualPremium" runat="server" Text="$1,209.60" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label46" runat="server" Text="Payment Form: " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblPaymentForm" runat="server" Text="Direct" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label45" runat="server" Text="Paid-to-Date: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblPaidToDate" runat="server" Text="01/02/2019" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label48" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label49" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label40" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label41" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label42" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label43" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label53" runat="server" Text="Servicing Agent: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblServiceAgent" runat="server" Text="JERRY ECK (A31800)" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label55" runat="server" Text="SA Phone: " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblSAPhone" runat="server" Text="(434) 808-2281" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label57" runat="server" Text="Servicing Agency: " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        &nbsp;<asp:Label ID="lblServiceAgency" runat="server" Text="INSPHERE (INS)	" CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label59" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label60" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label47" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label50" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label51" runat="server" Text=" " CssClass="bluetexttxt"></asp:Label>
                    </td>
                    <td align="left">
                        <asp:Label ID="Label52" runat="server" Text=" " CssClass="bluetexttxt" ></asp:Label>
                    </td>
                </tr>
             </table>
            <br />
             <table BORDER="0" CELLPADDING="1" CELLSPACING="0" width="100%">
	            <tr >
		            <td align="center" CssClass="bluetexttxt" >Check Requirements:</font></td>
	            </tr>
        
            </table>
            <asp:Button ID="btnEmail" runat="server" CssClass="clsButton" Text="Email New business Regarding Pending Requirements" OnClick="btnEmail_Click" />
&nbsp;<asp:Button ID="btnSource" runat="server" CssClass="clsButton" Text="Source Access" OnClick="btnSource_Click" />
&nbsp;<asp:Button ID="btnEMSI" runat="server" CssClass="clsButton" Text="      EMSI      " OnClick="btnEMSI_Click" />
              <br />
              <br />
             <table style="border-style:Solid;border-color:Gray;border-width: 1px;border-frame: box" BORDER="0" CELLPADDING="1" CELLSPACING="0" width="100%">
	            <tr bgcolor='#6699CC'>
		            <td align="center"><font color="white"><strong>Completed & Pending Requirements</strong></font></td>
	            </tr>
            </table>
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" Width="100%">
                <Columns>
                    <asp:BoundField DataField="OWNER_NAME" HeaderText="Name" ReadOnly="True" SortExpression="OWNER_NAME" >
                    <HeaderStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="REQ_DESCRIPTION" HeaderText="Requirement" SortExpression="REQ_DESCRIPTION" >
                    <HeaderStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="COMMENTS" HeaderText="Comments" SortExpression="COMMENTS" />
                    <asp:BoundField DataField="UND_FLAG" HeaderText="Met?" SortExpression="UND_FLAG" />
                    <asp:BoundField DataField="UND_DATE" HeaderText="Required Date" SortExpression="UND_DATE" />
                    <asp:BoundField DataField="UND_O_DATE" HeaderText="Received Date" SortExpression="UND_O_DATE" />
                </Columns>
            </asp:GridView>
            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:WRE_AgentConnectionString %>" SelectCommand="SELECT [INDIVIDUAL_FIRST] + ' ' +  [INDIVIDUAL_LAST] as OWNER_NAME, 
TRANSLATION.DESCRIPTION AS REQ_DESCRIPTION, [RECEIPT_FLAG] AS UND_FLAG, [COMMENT] AS COMMENTS, dbo.LPDATE_TO_STRDATE(ADD_DATE) AS UND_DATE, ADD_DATE, [MET_DATE] AS UND_O_DATE,  [RECEIPT_FLAG]  FROM [PENDING_REQUIREMENTS] LEFT OUTER JOIN TRANSLATION 
on TRANS_NAME = 'PENDING NB:REQUIREMENT TYPE' AND CODE = RECORD_TYPE WHERE (([COMPANY_CODE] = @COMPANY_CODE) AND ([POLICY_NUMBER] = @POLICY_NUMBER)) ORDER BY [ADD_DATE]
">
                <SelectParameters>
                    <asp:QueryStringParameter DefaultValue="07" Name="COMPANY_CODE" QueryStringField="COMPANY_CODE" />
                    <asp:QueryStringParameter DefaultValue="W000010138" Name="POLICY_NUMBER" QueryStringField="POLICY_NUMBER" />
                </SelectParameters>
            </asp:SqlDataSource>
              <br />
              <br />
            <table width="100%" class="auto-style1" >
                <tr style="width:100%">
                    <td align="left" 
                        style="font-family: arial, Helvetica, sans-serif; color: white" 
                        bgcolor="#0099FF" 
                        class="auto-style4"><b>(c) 2019 The Chesapeake Life Insurance Company, All rights reserved.</b></td>
                </tr> 
</table>
    </asp:Panel>
    </form>
    
</body>
</html>
