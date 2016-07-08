<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExportExcelPage.aspx.cs" Inherits="ExcelDemoProject.ExportExcelPage" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button ID="export_btn" runat="server" Text="匯出檔案" OnClick="export_btn_Click"/>
    </div>
    </form>
</body>
</html>
