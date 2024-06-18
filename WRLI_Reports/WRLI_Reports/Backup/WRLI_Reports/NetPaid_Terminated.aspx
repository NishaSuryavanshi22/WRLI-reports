<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="NetPaid_Terminated.aspx.cs" Inherits="WRLI_Reports._NetPaidTerminated" %>
 <%@ Register Namespace="AjaxControlToolkit" Assembly="AjaxControlToolkit" tagPrefix="ajax" %>
<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
    <style type="text/css">
        .style1
        {
            width: 533px;
        }
        .style2
        {
            width: 100px;
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


    <ajax:CalendarExtender ID="CalendarExtender1" runat="server" TargetControlID="TextBox1" PopupButtonID="Image1" Format ="MM/dd/yyyy">
    </ajax:CalendarExtender>
      <ajax:CalendarExtender ID="CalendarExtender2" runat="server" TargetControlID="TextBox2" PopupButtonID="Image2" Format ="MM/dd/yyyy">
    </ajax:CalendarExtender>

    <h3> <b>     Net Paid by Region for WRE</b>  </h3>
    
    <table width="100%">
       
        <tr>
            <td class="style2">Select Dates: </td>
            <td >
                <asp:TextBox ID="TextBox1" runat="server" ></asp:TextBox>
                <asp:Image ID="Image1" runat="server" ImageUrl="~/AjaxCalendarExtendar.gif"  /> to:
                <asp:TextBox ID="TextBox2" runat="server" >
                </asp:TextBox> <asp:Image ID="Image2" runat="server" ImageUrl="~/AjaxCalendarExtendar.gif" />
                
               
            </td>

            <td class="style2">
                Order Results by :
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlagent" runat="server"  OnPreRender="ddlagent_PreRender" >
                </asp:DropDownList>
                <asp:DropDownList ID="ddlSort" runat="server"  OnPreRender="ddlSort_PreRender">
                 <asp:ListItem>Descending</asp:ListItem>
                    <asp:ListItem>Ascending</asp:ListItem>
                </asp:DropDownList>
                 <asp:Button ID="Button4" runat="server" Text="Go!" OnClick="Button1_Click"/>
            </td>
            </tr>
        
        <tr>
            <td class="style2" >
                Select Type :
            </td>
            <td class="style1" >
                <asp:RadioButton ID="rdType"  runat="server" Text="All" ></asp:RadioButton>
                <asp:RadioButton ID="RadioButton1"  runat="server" Text="Med" ></asp:RadioButton>
                <asp:RadioButton ID="RadioButton2"  runat="server" Text="Non Med" >
                </asp:RadioButton>
                <asp:CheckBox ID="chkGoGreen" runat ="server" Text="GO Green" ForeColor="Green" />
                 </td>
        </tr>
        
           
</table>
    <table width="910px" id="tblgrid" runat="server">
        <tr>
            <td align="left" style="width:100%">
                <asp:Button ID="Button3" runat="server" Text="Print" OnClientClick="Print('dvgr');" />
                <asp:Button ID="Button2" runat="server" Text="Export To Excel" OnClick="Button2_Click" />
            </td>
        </tr>
        <tr>
            <td>
                <div style="border: thin groove #808080;width:910px; overflow: scroll;" runat="server"
                    id="dvgrid">
                    <div id="dvgr" style="width:100%">
                        <asp:GridView ID="GridView1" runat="server" CellPadding="4" EnableModelValidation="True"
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
</asp:Content>
