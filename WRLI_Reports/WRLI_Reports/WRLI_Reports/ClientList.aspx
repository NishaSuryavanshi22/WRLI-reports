<%@ Page Title="Home page" MasterPageFile="~/Site.Master"  Language="C#" AutoEventWireup="true" 
    CodeBehind="ClientList.aspx.cs" Inherits="WRLI_Reports.ClientList" %>

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

    <div style="margin: 0px auto 0px auto; width: 400px; ">
        <h2 >CLIENT LIST :</h2>
</div>
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
                        <asp:GridView ID="grdHandling" runat="server" CellPadding="4" EnableModelValidation="True"  RowStyle-Wrap="false"
                            ForeColor="#333333" GridLines="None" AllowPaging="true" PageSize="50" PagerSettings-Mode="NumericFirstLast" OnPageIndexChanging="grdHandling_PageIndexChanging" 
                            OnRowDataBound="grdHandling_RowDataBound" AutoGenerateColumns="true" OnSelectedIndexChanged="grdHandling_SelectedIndexChanged">
                            <AlternatingRowStyle BackColor="White" />
                            <EditRowStyle BackColor="#2461BF" />
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <RowStyle BackColor="#EFF3FB" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                            
                        </asp:GridView>
                       
                        <div>
                    Total Policy count :
                    <asp:Label ID="LBLPolicyCount" runat="server"></asp:Label>
                </div>
                       
                       

                    </div>
                </div>
         </td>
            </tr>
         </table>


 </asp:Content>