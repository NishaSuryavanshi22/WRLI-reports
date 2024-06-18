<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="WRLI_Reports.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">

        <div style="margin: 0px auto 0px auto; width: 400px;">
    <h2>Claims For Region :</h2>
</div>
        <div>
             <table width="910px" id="tblgrid" runat="server">
     <tr>
         <td align="right" style="width: 100%">
             <asp:Button ID="Button3" runat="server" Text="Print" OnClientClick="Print('tblgrid');" />
             <asp:Button ID="Button2" runat="server" Text="Export Report To Excel" OnClick="Button2_Click" />
            <asp:Button ID="Button1" runat="server" Text="Back" OnClick="Button2_Click" />

         </td>
     </tr>
     <tr>
         <td>
             <div style="border: thin groove #808080; width: 910px; overflow: scroll;" runat="server"
                 id="dvgrid">
                 <div id="dvgr" style="width: 100%">
                     <asp:GridView ID="grdHandling" runat="server" CellPadding="4" EnableModelValidation="True" RowStyle-Wrap="false"
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

                    


                 </div>
             </div>
         </td>
     </tr>
 </table>
        </div>
    </form>
</body>
</html>
