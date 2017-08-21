<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DashBoard.aspx.cs" Inherits="DashBoard" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <h1>Dox-Pro DashBoard</h1>
        <table>
            <tr>
                <td>
                    <asp:DataList runat="server" ID="data">
                    <HeaderStyle BackColor="#aaaadd"/>
                    <AlternatingItemStyle BackColor="Gainsboro"/>
                    <HeaderTemplate><b>Importer process report</b></HeaderTemplate>

                     <ItemTemplate>
                        <br />
                        <h2><%# ((DateTime)DataBinder.Eval(Container.DataItem, "Date")).ToString("dd/MM/yyyy") %></b><br /></h2>
                        <table>
                            <tr>
                                <td rowspan="3" style="padding-right: 10px"><b>Packing slips</b></td>
                                <td>Successfully imported:</td>
                                <td><%# DataBinder.Eval(Container.DataItem, "PsSucced", "{0:d}")%></td>
                                <td>(<%# DataBinder.Eval(Container.DataItem, "PsSuccedPercent", "{0:f2}")%>%)</td>
                            </tr>
                            <tr>
                                <td>Moved to manual queue:</td>
                                <td><%# DataBinder.Eval(Container.DataItem, "PsMovedToManual", "{0:d}")%></td>
                                <td>(<%# DataBinder.Eval(Container.DataItem, "PsMovedToManualPercent", "{0:f2}")%>%)</td>
                            </tr>
                            <tr>
                                <td>Errors:</td>
                                <td><%# DataBinder.Eval(Container.DataItem, "PsError", "{0:d}")%></td>
                                <td>(<%# DataBinder.Eval(Container.DataItem, "PsErrorPercent", "{0:f2}")%>%)</td>
                            </tr>
                            <tr>
                            <td colspan="4"><hr /></td>
                            </tr>
                            <tr>
                                <td rowspan="3" style="padding-right: 10px"><b>Invoices</b></td>
                                <td>Successfully imported:</td>
                                <td><%# DataBinder.Eval(Container.DataItem, "InvoceSSucced", "{0:d}")%></td>
                                <td>(<%# DataBinder.Eval(Container.DataItem, "InvoceSSuccedPercent", "{0:f2}")%>%)</td>
                            </tr>
                            <tr>
                                <td>Errors:</td>
                                <td><%# DataBinder.Eval(Container.DataItem, "InvoicesFailed", "{0:d}")%></td>
                                <td>(<%# DataBinder.Eval(Container.DataItem, "InvoicesFailedPercent", "{0:f2}")%>%)</td>
                            </tr>                
                        </table>
                        <br />
                        <br />
                     </ItemTemplate>
                    </asp:DataList>
                </td>
                <td valign="top" style="padding:10px ; border: 4px inset; background-color: Silver;">
                    <table>
                        <tr>
                            <td>Packing slips waiting to be processed:</td>
                            <td><% =DayActivity.PsWaitingInQueue.ToString() %></td>
                        </tr>
                        <tr>
                            <td>Packing slips waiting in manual queue:</td>
                            <td><% =DayActivity.PsTotalManual.ToString() %></td>
                        </tr>
                        <tr>
                            <td>Packing slips failed to be processed:</td>
                            <td><% =DayActivity.PsTotalError.ToString() %></td>
                        </tr>
                        <tr>
                            <td>Invoices waiting to be processed:</td>
                            <td><% =DayActivity.InvoicesTotalWaitingInQueue.ToString() %></td>
                        </tr>
                        <tr>
                            <td>Packing slips failed to be proceced:</td>
                            <td><% =DayActivity.InvoicesTotalFailed.ToString() %></td>
                        </tr>

                    </table>
                </td>
            </tr>
        </table>
 
    </div>
    </form>
</body>
</html>
