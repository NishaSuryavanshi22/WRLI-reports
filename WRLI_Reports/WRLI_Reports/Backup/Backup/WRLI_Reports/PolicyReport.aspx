<%@ Page Title="Policy Report" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="PolicyReport.aspx.cs" Inherits="WRLI_Reports._Default" %>
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
    function Print(dvgr) {
        var prntData = document.getElementById(dvgr);
        var prntWindow = window.open("", "Print", "width=20,height=20,left=0,top=0,toolbar=0,scrollbar=0,status=0");
       // prntWindow.document.write(prntData.InnerHTML);
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

    <h2>
        Policy Report
    </h2>
    
    <table width="100%">
        <tr>
            <td class="style2">
                Select Company :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlcompany" runat="server" AutoPostBack="True" >
                   </asp:DropDownList>
            </td>
        </tr>
        <!-- Agent name and code stored --->
        <tr  style="visibility:visible"  >
            <td class="style2">
                Select Agent :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlagent" runat="server"  OnPreRender="ddlagent_PreRender" >
                </asp:DropDownList>
            </td>
        </tr>
        <!-- Agent name display --->
        <tr style="visibility:hidden; display:none "  >
            <td class="style2"  >
                Select Agent List :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlagentNameList" runat="server"   >
                </asp:DropDownList>
            </td>
        </tr>
        

        <tr>
            <td class="style2">
                Select Region :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlregion" runat="server"  OnPreRender="ddlregion_PreRender" >
                
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td class="style2">
                Select State :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlstate" runat="server"  OnPreRender="ddlstate_PreRender">
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td class="style2">
                Select Data Type :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddldatatype" runat="server">

                    <asp:ListItem>SUBMITTED</asp:ListItem>
                    <asp:ListItem>PAID</asp:ListItem>
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td class="style2">
                Select Policy Status :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlpolicystatus" runat="server">
                    <asp:ListItem>ALL</asp:ListItem>
                    <asp:ListItem>ACTIVE</asp:ListItem>
                    <asp:ListItem>TERMINATED</asp:ListItem>
                    <asp:ListItem>SUSPENDED</asp:ListItem>
                    <asp:ListItem>PENDING</asp:ListItem>
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td class="style2">
                Select Policy Description :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlpolicydesc" runat="server"  OnPreRender="ddlpolicydesc_PreRender">
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td class="style2">
                Select Dates :
            </td>
            <td class="style1" colspan="4">
                <asp:TextBox ID="txtFrom" runat="server" Style="margin-right: 5px"></asp:TextBox>
                <asp:Image ID="imgfrom" runat="server" ImageUrl="~/AjaxCalendarExtendar.gif" Style="margin-right: 15px" />
                To : 
                <asp:TextBox ID="txtTo" runat="server" Style="margin-right: 8px"></asp:TextBox> <asp:Image ID="imgTo" runat="server" ImageUrl="~/AjaxCalendarExtendar.gif" Style="margin-right: 15px"/>
                <asp:Button ID="Button1" runat="server" Text="Go!" OnClick="Button1_Click"/>
            </td>
            </tr>
        
           
</table>
    <table width="910px" id="tblgrid" runat="server">
        <tr>
            <td align="right" style="width:100%">
                <asp:Label ID="Label1" runat="server" Text="Label" Visible="False"></asp:Label>
                <asp:Button ID="Button3" runat="server" Text="Print" OnClientClick="Print('dvgr');" />
                <asp:Button ID="Button2" runat="server" Text="Export To Excel" 
                    OnClick="Button2_Click" UseSubmitBehavior="False" />
            </td>
        </tr>
        <tr>
            <td>
                <div style="border: thin groove #808080;width:910px; overflow: scroll;" runat="server"
                    id="dvgrid">
                    <div id="dvgr" style="width:100%">
                        <asp:GridView ID="GridView1" runat="server" CellPadding="4" 
                            EnableModelValidation="True" RowStyle-Wrap="false"
                            SortExpression="POLICY_NUMBER" AllowSorting="false" AllowPaging="true" 
                            ForeColor="#333333" GridLines="None" 
                            onpageindexchanged="GridView1_PageIndexChanged" 
                            onpageindexchanging="GridView1_PageIndexChanging" 
                            onselectedindexchanging="GridView1_SelectedIndexChanging" PageSize="20">
                            <AlternatingRowStyle BackColor="White" />
                            <EditRowStyle BackColor="#2461BF" />
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Left" />
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
</asp:Content>
