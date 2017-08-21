<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>DMS DoxPro - Flextronics Packing Sleeps Manual Archiving Queue</title>
    <style type="text/css">
     html, body {
        height: 100%;
     }
     body {
          margin-top: 20px;
          padding: 0 3em;
          font: 12pt normal 'Myriad Pro', Arial, sans-serif;
          background: rgb(208,208,208); /* Old browsers */
          background: -moz-radial-gradient(right center, #fff, #bbb);
          background: -webkit-radial-gradient(right center, #fff, #bbb);
    }
    </style>
</head>
<body>
    <div>DMS DoxPro - Flextronics Packing Slips Manual Archiving Queue</div>
    <form id="form1" runat="server">
    <br />
    <asp:Label runat="server" Text="Branch: "></asp:Label>
    <asp:DropDownList ID="Factory" runat="server" AutoPostBack="True" OnSelectedIndexChanged="Factory_OnSelectedIndexChanged">
        <asp:ListItem Selected="True">מגדל העמק</asp:ListItem>
        <asp:ListItem>אופקים</asp:ListItem>
    </asp:DropDownList>
    <br />
    <br />
    <div>
    <asp:Label ID="Label1" runat="server" Text="Total Files: "></asp:Label>
        <asp:Literal ID="Total" runat="server"></asp:Literal>
        <br />
    <div></div>
    <table>
    <tr>
        <td valign="top">
            <asp:ListBox ID="filesInQueue" Rows="40" runat="server"  AutoPostBack="true" style="height:700px;width:240px" 
                onselectedindexchanged="filesInQueue_SelectedIndexChanged"></asp:ListBox>
        </td>
        <td valign="top">
            <div id="pdfframe" style="height:700px;width:900px" runat="server" />
        </td>
        <td valign="top">
            <table>
            <tr>
                <td>Customer:</td>
                <td><asp:TextBox ID="customer" runat="server" /></td>
            </tr>
            <tr>
                <td>Packing sleep number:</td>
                <td><asp:TextBox ID="psNumber" runat="server" /></td>
            </tr>
            </table><br />
             
            <asp:Button ID="submitBarCode" runat="server" onclick="submitBarCode_Click" Text="Archive" />
                <br />
                <br />
            <asp:Button ID="Remove" runat="server" onclick="Remove_Click" Text="Delete file" />
                <br />
                <br />
            <asp:Button ID="manualQueueButton" runat="server" onclick="manualQueueButton_Click" Text="Move to manual queue" />
            <br />
        </td>
    </tr>
    </table>
    
    </div>
    </form>
</body>
</html>
