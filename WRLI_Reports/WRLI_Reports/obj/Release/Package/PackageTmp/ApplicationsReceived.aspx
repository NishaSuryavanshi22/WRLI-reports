<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="ApplicationsReceived.aspx.cs" Inherits="WRLI_Reports.ApplicationsReceived" %>

<%@ Register Namespace="AjaxControlToolkit" Assembly="AjaxControlToolkit" TagPrefix="ajax" %>
<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
    <style type="text/css">
        <b href="ApplicationsReceived.aspx" > ApplicationsReceived /b >
        .style1 {
            width: 633px;
        }

        .style2 {
            width: 193px;
        }
    </style>
    <a href="ApplicationsReceived.aspx">ApplicationsReceived.aspx</a>
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <script language="javascript">
    function Print(dvgr) {
        // var prntData = document.getElementById('<%= tblgrid.ClientID %>');
        var prntData = document.getElementById(dvgr);
        var prntWindow = window.open("", "Print", "width=20,height=20,left=0,top=0,toolbar=0,scrollbar=0,status=0");
        prntWindow.document.write(prntData.outerHTML);
        prntWindow.document.close();
        prntWindow.focus();
        prntWindow.print();
        prntWindow.close();
    }

    </script>

    <ajax:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server">
    </ajax:ToolkitScriptManager>




    <h2 style="width: 100%" id="headingInt">Application Received
    </h2>
    <br />
    <br />
    <div width="100%" id="tblHandletime">
        <table width="100%" id="tblHandle">

            <tr>
                <td style="width: 500px"><b>This page allows you to see if an application has been received in the Home Office. That's right, within an hour of an application being received in the Home Office, you will be able to check the Web Site to confirm that it has been received. This includes applications that have been mailed or Faxed. Applications will be listed here if they have been received within the last four days. 
       PLEASE NOTE: Within 24 to 48 hours, the application will also be available on the "My Policies" page. 
       Reminder: If you are faxing an application, be sure to always use the Fax Cover sheet with the application.  </b></td>
            </tr>
            <tr>
                <td align="left" style="width: 100%"></td>

                <td width="30px">Select Agent :
                </td>
                <td width="30px">
                    <!-- Agent name display --->
                    <asp:DropDownList ID="ddlagent" runat="server" Width="150px">
                    </asp:DropDownList>
                </td>
                <td class="style1" width="30px">
                    <asp:Button ID="Button4" runat="server" Text=" GO Refresh" OnClick="Button4_Click " />

                </td>
            </tr>


        </table>
        <table width="210px" id="tblgrid" runat="server">
            <tr>
                <td align="right" style="width: 100%" colspan="0">
                    <asp:Button ID="Button3" runat="server" Text="Print" OnClientClick="Print('dvgr');" />
                    <asp:Button ID="Button2" runat="server" Text="Export Report To Excel" OnClick="Button2_Click" />
                </td>
            </tr>
            <tr>
                <td class="style2" colspan="0"></td>
            </tr>

            <tr>
                <td colspan="0">

                    <div style="border: thin groove #808080; width: 910px; overflow: scroll;" runat="server"
                        id="dvgrid">
                        <div id="dvgr" style="width: 100%">
                            <asp:GridView ID="grInterviewsByRegion" runat="server" CellPadding="4" EnableModelValidation="True" RowStyle-Wrap="false"
                                ForeColor="#333333" GridLines="Both" Visible="true">
                                <AlternatingRowStyle BackColor="White" />
                                <EditRowStyle BackColor="#2461BF" />
                                <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#EFF3FB" />
                                <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </div>

                    <div>
                        <b>
                            <asp:Label ID="lblcount" runat="server" Text=""></asp:Label>
                        </b>
                    </div>
                </td>
            </tr>
        </table>
    </div>
</asp:Content>
