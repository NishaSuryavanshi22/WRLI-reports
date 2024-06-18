﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Ratedforagent.aspx.cs" Inherits="WRLI_Reports.Ratedforagent" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
         <h2  style="width:100%">
  Rated for Agent :
     <asp:Label ID="Agent_Number" colour="red" runat="server" Text=""></asp:Label> 
    </h2>
    <form id="form1" runat="server">
        <div>
                                <table width="910px" id="tblgrid" runat="server">
<tr>
    <td align="right" style="width: 100%">
        <asp:Button ID="Button3" runat="server" Text="Print" OnClientClick="Print('tblgrid');" />
        <asp:Button ID="Button2" runat="server" Text="Export Report To Excel" OnClick="Button2_Click" />
       <asp:Button ID="Button1" runat="server" Text="Back" OnClick="Button1_Click" />

    </td>
</tr>
<tr>
    <td>
        <div style="border: thin groove #808080; width: 910px; overflow: scroll;" runat="server"
            id="dvgrid">
            <div id="dvgr" style="width: 200%">
                <asp:GridView ID="grdratedforagent" onrowdatabound="grdratedforagent_RowDataBound" runat="server" CellPadding="4" EnableModelValidation="True" RowStyle-Wrap="false"
                    ForeColor="#333333" GridLines="None" AllowPaging="true" PageSize="50"  
                   AutoGenerateColumns="true" >
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

             <div>
    <b>
    <asp:Label ID="lblcount" runat="server" Text=""></asp:Label> </b>
</div>
       
        </div>
    </form>
</body>
</html>
