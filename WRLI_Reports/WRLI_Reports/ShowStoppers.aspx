<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="ShowStoppers.aspx.cs" Inherits="WRLI_Reports.ShowStoppers" %>
 <%@ Register Namespace="AjaxControlToolkit" Assembly="AjaxControlToolkit" tagPrefix="ajax" %>
<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
    <style type="text/css"><b href="ShowStoppers.aspx">ShowStoppers.aspx</b>
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
<script  type="text/javascript">
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


    

    <h2  style="width:100%" id="headingInt" >
        Show Stoppers
    </h2>
    <br /><br />
     <div width="100%" id="tblHandletime">
    <table width="100%" id="tblHandle">
       
        <tr>
            <td class="style2">
                Select Agent :
            </td>
            <td  >
                <!-- Agent name display --->
                <asp:DropDownList ID="ddlagent" runat="server" Width= "200px"  >
                </asp:DropDownList>
            </td>
            <td class="style1" >
                <asp:CheckBox ID="chkInterview" runat ="server" Text="Include Completed Showstoppers " ForeColor="Green" />
                 </td>
        </tr>
            <!-- Agent name display --->
            <tr><asp:Button ID="Button1" runat="server" Text=" Refresh " OnClick="Button1_Click"/></tr>
        
</table>
    <table width="910px" id="tblgrid" runat="server">
        <tr>
            <td align="right" style="width:100%">
                <asp:Button ID="Button3" runat="server" Text="Print" OnClientClick="Print('dvgr');" />
                <asp:Button ID="Button2" runat="server" Text="Export Report To Excel" OnClick="Button2_Click" />
            </td>
        </tr>
        <tr> 
        <td class="style2">
               
            </td>
        </tr>

        <tr>
            <td>
            
                <div style="border: thin groove #808080;width:910px; overflow: scroll;" runat="server"
                    id="dvgrid">
                    <div id="dvgr" style="width:100%">
                        <asp:GridView ID="grInterviewsByRegion" runat="server" CellPadding="4" EnableModelValidation="True"  RowStyle-Wrap="false"
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
                    <asp:Label ID="lblcount" runat="server" Text=""></asp:Label> </b>
                </div>
            </td>
        </tr>
    </table>
    </div>
</asp:Content>
