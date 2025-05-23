<%@ Page Language="vb" AutoEventWireup="false" MasterPageFile="../master/MasterMain.master" CodeBehind="DirectPayments.aspx.vb" Inherits="PrePress.DirectPayments" title="GEMS Direct Payments" %>
 <%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Register TagPrefix="it" TagName="PageMessage" src="~/usercontrols/PageMessage.ascx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
     <script type="text/javascript">

         var popUp;
         function PopUpShowing(sender, eventArgs) {
             popUp = eventArgs.get_popUp();
             var gridWidth = sender.get_element().offsetWidth;
             var gridHeight = sender.get_element().offsetHeight;
             var popUpWidth = popUp.style.width.substr(0, popUp.style.width.indexOf("px"));
             var popUpHeight = popUp.style.height.substr(0, popUp.style.height.indexOf("px"));
             popUp.style.left = ((gridWidth - popUpWidth) / 2 + sender.get_element().offsetLeft).toString() + "px";
             popUp.style.top = ((gridHeight - popUpHeight) / 2 + sender.get_element().offsetTop).toString() + "px";
         }

  </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
     <it:PageMessage ID="PageMessage" runat="Server" />

     <table id="Table1" border="0" cellpadding="5" cellspacing="5" runat="server">
        <tr id="trSearch" runat="server">
             <td style="width:100%" align='left'>
             <telerik:RadDatePicker AutoPostBack="false" ID="StartDT" Runat="server"  DateInput-EmptyMessage="Start Date" MinDate="1/1/1900" Width="100px" Culture="English (United States)">
                <Calendar ID="Calendar1" runat="server" UseRowHeadersAsSelectors="False"  ViewSelectorText="x">
                </Calendar>
                <DatePopupButton ImageUrl="" HoverImageUrl="">
                </DatePopupButton>
                <DateInput ID="DateInput1" runat="server" DisplayDateFormat="M/d/yyyy" DateFormat="M/d/yyyy">
                </DateInput>
            </telerik:RadDatePicker>
      &nbsp;To&nbsp;
       <telerik:RadDatePicker AutoPostBack="false" ID="EndDT" Runat="server"  DateInput-EmptyMessage="End Date" MinDate="1/1/1900" Width="100px" Culture="English (United States)">
                <Calendar ID="Calendar2" runat="server" UseRowHeadersAsSelectors="False" UseColumnHeadersAsSelectors="False" ViewSelectorText="x">
                </Calendar>
                <DatePopupButton ImageUrl="" HoverImageUrl="">
                </DatePopupButton>
                <DateInput ID="DateInput2" runat="server" DisplayDateFormat="M/d/yyyy" DateFormat="M/d/yyyy">
                </DateInput>
            </telerik:RadDatePicker>
      &nbsp;
        <telerik:RadButton runat="server" Text="Search" ID="btnSearchFilters" ></telerik:RadButton>
          &nbsp;
        <telerik:RadButton ID="btnExport" Text="Export to CSV" runat="server"></telerik:RadButton>
              </td>
        </tr>
    </table>

      <telerik:RadGrid ID="radDirectPayments" runat="server" AllowSorting="True" Skin="Default"
        AutoGenerateColumns="False" EnableLinqExpressions="false" GridLines="None" ShowStatusBar="true"
        AllowMultiRowEdit="false" AllowMultiRowSelection="true" Width="100%" ShowFooter="true">      
        <ClientSettings AllowDragToGroup="false" AllowColumnsReorder="false" ReorderColumnsOnClient="false" AllowRowsDragDrop="True"
            ColumnsReorderMethod="Reorder">         
            <Selecting AllowRowSelect="true" EnableDragToSelectRows="false" />           
            <Scrolling AllowScroll="false" SaveScrollPosition="false" />
            <Resizing AllowColumnResize="False" />    
             <ClientEvents OnPopUpShowing="PopUpShowing" />        
        </ClientSettings>    
        <FilterMenu EnableTheming="True">
            <CollapseAnimation Type="OutQuint" Duration="200"></CollapseAnimation>
        </FilterMenu>
        <%-- Master Table--%>
        <MasterTableView TableLayout="auto" EditMode="PopUp" AutoGenerateColumns="false"
            AllowSorting="true" CommandItemDisplay="top" CellSpacing="0" DataKeyNames="PaymentLogId">
            <SortExpressions>
                <telerik:GridSortExpression FieldName="CreatedOn" SortOrder="Ascending" />
            </SortExpressions>
            <%-- Top Command Bar --%>           
            <EditFormSettings UserControlName="../usercontrols/ManageDirectPayments.ascx" EditFormType="WebUserControl">               
                 <PopUpSettings Width="800px" Modal="true" />
            </EditFormSettings>           
            <CommandItemSettings ExportToPdfText="Export to PDF" AddNewRecordText="Add"></CommandItemSettings>
            <RowIndicatorColumn FilterControlAltText="Filter RowIndicator column">
            </RowIndicatorColumn>
            <ExpandCollapseColumn FilterControlAltText="Filter ExpandColumn column">
            </ExpandCollapseColumn>
            <Columns>
                <telerik:GridBoundColumn DataField="PaymentLogId" Display="False" Visible="True" ReadOnly="false"
                    UniqueName="PaymentLogId" HeaderText="PaymentLogId" AllowFiltering="false" Groupable="false">
                </telerik:GridBoundColumn>
             
              <telerik:GridBoundColumn DataField="CurrencyId" Display="False" Visible="True" ReadOnly="false"
                    UniqueName="CurrencyId" HeaderText="CurrencyId" AllowFiltering="false" Groupable="false">
                </telerik:GridBoundColumn>

              <telerik:GridBoundColumn DataField="Freelancer" SortExpression="Freelancer"
                    Display="true" Visible="true" ReadOnly="false" UniqueName="Freelancer"
                    HeaderText="Freelancer" AllowFiltering="true" Groupable="false">
                </telerik:GridBoundColumn>

                <telerik:GridBoundColumn DataField="Job" SortExpression="Job"
                    Display="true" Visible="true" ReadOnly="false" UniqueName="Job"
                    HeaderText="Job Number" AllowFiltering="true" Groupable="false">
                </telerik:GridBoundColumn>
                
                  <telerik:GridBoundColumn DataField="BatchNumber" SortExpression="BatchNumber"
                    Display="true" Visible="true" ReadOnly="false" UniqueName="BatchNumber"
                    HeaderText="Batch Number" AllowFiltering="true" Groupable="false">
                </telerik:GridBoundColumn>     
                            
                  <telerik:GridBoundColumn DataField="EditorialService" SortExpression="EditorialService"
                    Display="true" Visible="true" ReadOnly="false" UniqueName="EditorialService"
                    HeaderText="Editorial Service" AllowFiltering="true" Groupable="false">
                </telerik:GridBoundColumn>     

                <telerik:GridBoundColumn DataField="UnitType" SortExpression="UnitType"
                    Display="true" Visible="true" ReadOnly="false" UniqueName="UnitType"
                    HeaderText="Unit Type" AllowFiltering="true" Groupable="false">
                </telerik:GridBoundColumn>     

                 <telerik:GridBoundColumn DataField="JobType" SortExpression="JobType"
                    Display="true" Visible="true" ReadOnly="false" UniqueName="JobType"
                    HeaderText="Job Type" AllowFiltering="true" Groupable="false">
                </telerik:GridBoundColumn>     
                 
                <telerik:GridBoundColumn DataField="IsLumpsum" SortExpression="IsLumpsum"
                    Display="true" Visible="true" ReadOnly="false" UniqueName="IsLumpsum"
                    HeaderText="Is Lumpsum" AllowFiltering="true" Groupable="false">
                </telerik:GridBoundColumn> 

                 <telerik:GridBoundColumn DataField="UnitCount" SortExpression="UnitCount"
                    Display="true" Visible="true" ReadOnly="false" UniqueName="UnitCount"
                    HeaderText="Unit Count" AllowFiltering="true" Groupable="false">
                </telerik:GridBoundColumn> 

                  <telerik:GridBoundColumn DataField="PricePerPageDisplay" SortExpression="PricePerPage"
                    Display="true" Visible="true" ReadOnly="false" UniqueName="PricePerPageDisplay"
                    HeaderText="Price Per Page" AllowFiltering="true" Groupable="false">
                </telerik:GridBoundColumn> 

                  <telerik:GridBoundColumn DataField="TotalPriceDisplay" SortExpression="TotalPrice"
                    Display="true" Visible="true" ReadOnly="false" UniqueName="TotalPriceDisplay"
                    HeaderText="Total Price" AllowFiltering="true" Groupable="false">
                </telerik:GridBoundColumn> 

               <%-- <telerik:GridTemplateColumn Groupable="false" HeaderStyle-Width="150px" HeaderStyle-Wrap="false"
                    ItemStyle-Wrap="false" SortExpression="TotalPrice" HeaderText="Total Price">
                    <ItemTemplate>
                        <%# Eval("CurrencyDisplay")%><%# Eval("TotalPrice", "{0:##,##0.0.0#}")%>
                    </ItemTemplate>                  
                    <HeaderStyle Wrap="False" Width="150px"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </telerik:GridTemplateColumn>--%>

                 <telerik:GridBoundColumn DataField="isPaid" SortExpression="IsPaidTF" Display="true"
                    Visible="true" ReadOnly="false" UniqueName="IsPaidTF" HeaderText="Paid" AllowFiltering="false"
                    Groupable="false">
                </telerik:GridBoundColumn>

                 <telerik:GridBoundColumn DataField="PaymentDT" SortExpression="PaymentDT"
                    DataType="System.DateTime" DataFormatString="{0:M/d/yyyy}" Display="true" Visible="true"
                    ReadOnly="false" UniqueName="PaymentDT" HeaderText="Payment Date"
                    AllowFiltering="false" Groupable="false">
                </telerik:GridBoundColumn>
                          
                  <telerik:GridEditCommandColumn UpdateText="Update" HeaderStyle-Width="20px" UniqueName="EditCommandColumn" CancelText="Cancel" EditText="Edit"    
                Reorderable="false">
            </telerik:GridEditCommandColumn>
            <telerik:GridButtonColumn CommandName="Delete" Text="Delete" HeaderStyle-Width="20px" UniqueName="DeleteColumn" ConfirmText="This record will be permanently deleted. Do you wish to continue?" > 
           </telerik:GridButtonColumn>
                <%--   <telerik:GridButtonColumn CommandName="View" Text="View" HeaderStyle-Width="20px" UniqueName="ViewColumn" > </telerik:GridButtonColumn>--%>
            </Columns>
        </MasterTableView>
        <HeaderContextMenu CssClass="GridContextMenu GridContextMenu_Default">
        </HeaderContextMenu>
        <PagerStyle Mode="Slider" />
        <FilterMenu Skin="Default">
            <CollapseAnimation Type="OutQuint" Duration="200"></CollapseAnimation>
        </FilterMenu>
    </telerik:RadGrid>

     <telerik:RadWindowManager ID="RadWindowManager1" Style="z-index: 7200"   ShowContentDuringLoad="false" VisibleStatusbar="false"  Modal="true"  Behaviors="Close, Move, Resize" Height="550px" Width="550px"
        ReloadOnShow="true" runat="server" Skin="Default" EnableShadow="true"  >
    </telerik:RadWindowManager>

    <script>
        function fnClose() {
            var manager = GetRadWindowManager();
            var window1 = manager.GetWindowById(radviewwin);
            window1.close();
            return false;
        }

    </script>

</asp:Content>