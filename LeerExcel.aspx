<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="LeerExcel.aspx.cs" Inherits="PruebasExcel.LeerExcel" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
              <asp:FileUpload ID="fuExcel" runat="server" />
  <asp:TextBox ID="txtSheetNumber" runat="server" placeholder="Número de hoja"></asp:TextBox>
  <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" />
  <br />
  <asp:Literal ID="ltExcelTable" runat="server"></asp:Literal>
        </div>
    </form>
</body>
</html>
