<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="PerformanceReport.aspx.cs" Inherits="WRLI_Reports.PerformanceReport" %>
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

   <%-- <script type="text/javascript">
        window.history.replaceState({}, document.title, window.location.pathname);
</script>--%>

<script language="javascript" type="text/javascript">
    function Print(tblContainer) {
        var prntData = document.getElementById('<%= tblContainer.ClientID %>');
        var prntWindow = window.open("Print Report", "Print", "width=700,height=700,left=0,top=0,toolbar=0,scrollbar=1,status=0");
        prntWindow.document.write(prntData.outerHTML);
        prntWindow.document.close();
        prntWindow.focus();
        prntWindow.print();
        prntWindow.close();
    }

    function Print_new(documentId) {
       // debugger;
       // window.parent.main.focus();
        window.print();
    }

    function OnClientDateSelectionChanged() {
    }

    function Print_newaa(documentId) {

        //Wait until PDF is ready to print  
        
        if (typeof document.getElementById(documentId).print == 'undefined') {

            setTimeout(function () { printDocument(documentId); }, 1000);

        } else {

            //var x = document.getElementById(documentId);
            var x = document.getElementById('<%= tblContainer.ClientID %>').outerHTML;
            x.print();
        }
    }

    function OnPageInit(sname) {

        //alert(sname);
        
    }



</script>

    <ajax:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server">
    </ajax:ToolkitScriptManager>


    <ajax:CalendarExtender ID="CalendarExtender1" runat="server" TargetControlID="TextBox1" PopupButtonID="Image1" Format ="MM/dd/yyyy">
    </ajax:CalendarExtender>
      <ajax:CalendarExtender ID="CalendarExtender2" runat="server" TargetControlID="TextBox2" PopupButtonID="Image2" Format ="MM/dd/yyyy">
    </ajax:CalendarExtender>

    <h2> <b>     Performance Report</b>  </h2>
    <br />
    
    <table width="95%" id ="tblContainer"  runat="server" align="center" >
       
        
            <tr >
            <td class="style2" colspan="6">
                
                <asp:Label runat="server" ID ="lblAgent" Text="Select Agent" ></asp:Label>
                <asp:DropDownList ID="ddlListAgent" runat="server"  OnPreRender="ddlagent_PreRender">
                </asp:DropDownList>
            </td>
            </tr>
            <tr >
            <td class="style2" colspan="6">
                <asp:CheckBox ID="chkBoxOverride" runat ="server" Text="Allow date override" ForeColor="Green" />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Label runat="server" ID ="lblStart" Text="Start Date" ></asp:Label>
                <asp:TextBox ID="TextBox1" runat="server" ></asp:TextBox>
                <asp:Image ID="Image1" runat="server" ImageUrl="~/AjaxCalendarExtendar.gif"  /> To:
                <asp:TextBox ID="TextBox2" runat="server" >
                </asp:TextBox> <asp:Image ID="Image2" runat="server" ImageUrl="~/AjaxCalendarExtendar.gif" /> ( 15 Months Performance from start date)
                <asp:Calendar ID="cdStartdate" runat="server" Caption="Start Date" ShowTitle="true" ShowGridLines ="true"  Visible="false" ToolTip ="Select start date"
                 OnSelectionChanged="OnClientDateSelectionChanged"  > </asp:Calendar>
                
                
               
            </td>
            
             </tr>
             <tr>
            <td class="style2" colspan="6">
                
                <asp:Label runat="server" ID ="lblReport" Text="Select Report" ></asp:Label>
                <asp:DropDownList ID="ddlBussReport" runat="server"  OnPreRender="ddlBussReport_PreRender">
                </asp:DropDownList>
            </td>
            </tr>

             <tr>
            <td class="style2" colspan="6">
                
                <asp:Label runat="server" ID ="lblRegion" Text="Select Region" ></asp:Label>
                <asp:DropDownList ID="ddlRegion" runat="server"  OnPreRender="ddlRegion_PreRender">
                </asp:DropDownList>
                <asp:CheckBox ID="chkCompany" runat ="server" Text="Get Company Overview" ForeColor="Green" AutoPostBack="true" OnCheckedChanged="DiableIndividualAgent"   />
            </td>
            </tr>
            
        
             <tr>
            <td class="style1" colspan="5">
                
            <asp:RadioButton ID="RadioButton1"  runat="server" Text="Get Results by Hierarchy Agent "  GroupName="rbHier" ></asp:RadioButton>
            <asp:RadioButton ID="RadioButton2"  runat="server" Text="Get Results by Individual Agent "  GroupName="rbHier" >
            </asp:RadioButton>
            
            <br /><br />
            </td>
        </tr>

            <tr>
            <td colspan="5">

                <asp:CheckBox ID="chkGoGreen" Font-Bold="true" runat ="server" Text="GO Green" ForeColor="Green" />
                <asp:Button ID="btnGo" runat="server" Text="GO" OnClick="Go_Click" />
            </td>
             </tr>
           

        <tr>
            <td align="left" style="width:100%">
                <asp:Button ID="Button3" runat="server" Text="Print" OnClientClick="Print_new('dvgr');" />
                <asp:Button ID="Button2" runat="server" Text="Export To Excel" OnClick="Button2_Click" />
            </td>
        </tr>
        <tr>
            <td>
                <div style="border: thin groove #808080;width:100%; overflow: scroll;" runat="server"
                    id="dvgrid">
                    <div id="dvgr" style="width:100%;height:100%">
                        <asp:GridView ID="GridView1" runat="server" CellPadding="4" EnableModelValidation="True"  EmptyDataRowStyle-BackColor ="Red" 
                        RowStyle-BackColor="Yellow" FooterStyle-BackColor="Green"
                            ForeColor="#333333" GridLines="Both" Width="99%" Height="99%">
                            <AlternatingRowStyle BackColor="White" />
                            <EditRowStyle BackColor="#2461BF"   /> 
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <RowStyle BackColor="#EFF3FB" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                        </asp:GridView>
                                            
                    </div>
                </div>
                <div>
                    
                    <asp:Label ID="lblcount" runat="server" Text="0" Font-Bold="true"></asp:Label>
                </div>
                <asp:HiddenField ID ="hdnGridVW" Value="" runat="server" />
            </td>
        </tr>
    </table>
</asp:Content>
