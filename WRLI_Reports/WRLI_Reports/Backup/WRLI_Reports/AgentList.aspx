<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="AgentList.aspx.cs" Inherits="WRLI_Reports.AgentList" %>
 <%@ Register Namespace="AjaxControlToolkit" Assembly="AjaxControlToolkit" tagPrefix="ajax" %>
<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
    <style type="text/css">
        .style1
        {
            width: 633px;
        }
        .style2
        {
            width: 193px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
<script language="javascript">
    function Print(tblgrid) {
        var prntData = document.getElementById('<%= tblgrid.ClientID %>');
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


    <h2>
        Agent List
    </h2>
     <div id="tblHandletime">
    <table width="100%" id="tblHandle">
        <tr>
            <td class="style2">
                Select Company :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlHandcompany" runat="server">
                   </asp:DropDownList>
            </td>
        </tr>
        
       
        <tr>
            <td class="style2">
                Select Policy Status :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlHandpolicystatus" runat="server">
                    <asp:ListItem>ALL</asp:ListItem>
                    <asp:ListItem>ACTIVE</asp:ListItem>
                    <asp:ListItem>TERMINATED</asp:ListItem>
                    
                </asp:DropDownList>
            </td>
        </tr>

        <tr>
            <td class="style1" colspan="0">
                <asp:Button ID="btnGo" runat="server" Text="Go!" OnClick="Button1_Click"/>
            </td>
            </tr>
           
           
</table>
 <b> Select the Options you wish to include in this report and click the Go! button </b>        
    <table width="910px" id="tblgrid" runat="server">
        <tr>
            <td align="right" style="width:100%">
                <asp:Button ID="Button3" runat="server" Text="Print" OnClientClick="Print('tblgrid');" />
                <asp:Button ID="Button2" runat="server" Text="Export Report To Excel" OnClick="Button2_Click" />
            </td>
        </tr>
        <tr>
            <td>
                <div style="border: thin groove #808080;width:910px; overflow: scroll;" runat="server"
                    id="dvgrid">
                    <div id="dvgr" style="width:100%">
                        <asp:GridView ID="grdAgentList" runat="server" CellPadding="4" EnableModelValidation="True"  RowStyle-Wrap="false" EnableSortingAndPagingCallbacks ="true"
                            ForeColor="#333333" GridLines="None">
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
                    Total Policy count :
                    <asp:Label ID="lblcount" runat="server" Text=""></asp:Label>
                </div>
            </td>
        </tr>
    </table>
    </div>
</asp:Content>
