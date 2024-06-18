<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="PerformanceReport_AllRegion.aspx.cs" Inherits="WRLI_Reports.PerformanceReport_AllRegion" %>
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
<script language="javascript" type="text/javascript">
    function Print(dvgr) {
        var prntData = document.getElementById(dvgr);
        var prntWindow = window.open("", "Print", "width=20,height=20,left=0,top=0,toolbar=0,scrollbar=0,status=0");
       // prntWindow.document.write(prntData.InnerHTML);
        prntWindow.document.close();
        prntWindow.focus();
        prntWindow.print();
        prntWindow.close();
    }

    
    function OnClientDateSelectionChanged() {
        
    }


</script>


    <h2> <b>     Performance Report</b>  </h2>
    <br />
    
    <table width="100%">
       
        
            <tr >
            <td class="style2" >
                
                <asp:Label runat="server" ID ="lblAgent" Text="Hierarchy is hierarchy" ></asp:Label>
             </td>
            </tr>
            
             <tr>
            <td class="style2" >
                 <asp:Label runat="server" ID ="lblReport" Text="New Business Performance by region" ></asp:Label>
                
            </td>
            </tr>

             <tr>
            <td class="style2" >
                
                <asp:Label runat="server" ID ="lblRegion" Text="Date Range:"  Font-Bold="true"></asp:Label>
               
               
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
                
               
            </td>
        </tr>

        <tr> <td >
            <asp:Label runat="server" ID ="lblNoRecords" Text="No Records to show" ></asp:Label>
            </td> </tr>
        
         <tr>
            <td >
             <asp:Label runat="server" ID ="lblSplitsCt" Text="Note: Splits are counted by whole number" ></asp:Label>
                <asp:Button ID="btnGo" runat="server" Text="Back" OnClick="Back_Click" />
            </td> </tr>
         </table>
</asp:Content>
