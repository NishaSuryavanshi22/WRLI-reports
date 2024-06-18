<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="Pending_PoliciesByAgent.aspx.cs" Inherits="WRLI_Reports.Pending_PoliciesByAgent" %>
 <%@ Register Namespace="AjaxControlToolkit" Assembly="AjaxControlToolkit" tagPrefix="ajax" %>
<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
    <style type="text/css"><b href="Pending_PoliciesByAgent.aspx">Paid, Placement and Persistency Report </b>
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
<script type="text/javascript" >
    function Print(dvgr) {
        // var prntData = document.getElementById('<%= tblgrid.ClientID %>');
        var prntData = document.getElementById(dvgr);
        if (prntData == null)
            return;
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


    <ajax:CalendarExtender ID="CalendarExtender1" runat="server" TargetControlID="txtFrom" PopupButtonID="imgfrom" Format ="MM/dd/yyyy">
    </ajax:CalendarExtender>
      <ajax:CalendarExtender ID="CalendarExtender2" runat="server" TargetControlID="txtTo" PopupButtonID="imgTo" Format ="MM/dd/yyyy">
    </ajax:CalendarExtender>

    <h2  style="width:100%">
        Policies by Agent
    </h2>
    <br /><br />
     <div width="100%" id="tblHandletime">
    <table width="100%" id="tblHandle">
       
        <tr>

        <tr>
            <td class="style2">
                Select Agent :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlAgentList" runat="server"  OnPreRender="ddlagent_PreRender">
                </asp:DropDownList>
            </td>
        </tr>

            <td class="style2">
                Select Dates :
            </td>
            <td class="style1" >
                <asp:TextBox ID="txtFrom" runat="server" Style="margin-right: 5px"></asp:TextBox>
                <asp:Image ID="imgfrom" runat="server" ImageUrl="~/AjaxCalendarExtendar.gif" Style="margin-right: 15px" />
                to:
                <asp:TextBox ID="txtTo" runat="server" Style="margin-right: 8px"  ></asp:TextBox> 
                <asp:Image ID="imgTo" runat="server" ImageUrl="~/AjaxCalendarExtendar.gif" Style="margin-right: 15px"  />
<%--                <asp:Button ID="GO" runat="server" Text="Go!" OnClick="Button1_Click"/>--%>
                
            </td>
            </tr>
        <tr>
            
            <td class="style1" >
                
                <asp:CheckBox ID="chkGoGreen" runat ="server" Text="Go Green" ForeColor="Green" />
                 </td>
        </tr>
           
</table>
    <table width="910px" id="tblgrid" runat="server">
        <tr>
            <td align="right" style="width:100%">
            <asp:Button ID="All" runat="server" Text="ALL" OnClick="All_Click" Visible="true" />
            <asp:Button ID="Active" runat="server" Text="Active" OnClick="Active_Click" Visible ="true"/>
            <asp:Button ID="Pending" runat="server" Text="Pending" OnClick="Pending_Click"/>
            <asp:Button ID="Terminated" runat="server" Text="Terminated" OnClick="Terminated_Click"/>

                <asp:Button ID="Button3" runat="server" Text="Print" OnClientClick="Print('dvgr');" />
                <asp:Button ID="Button2" runat="server" Text="Export Report To Excel" OnClick="Button2_Click" />
            </td>
        </tr>
        <tr>
            <td>
            
                <div style="border: thin groove #808080;width:910px; overflow: scroll;" runat="server"
                    id="dvgrid">
                    <div id="dvgr" style="width:100%">
                        <asp:GridView ID="grdSubmittedReport" runat="server" CellPadding="4" EnableModelValidation="True"  RowStyle-Wrap="false"
                            ForeColor="#333333" GridLines="Both" Visible="true" Font-Bold="FALSE" EnableSortingAndPagingCallbacks ="true">
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
