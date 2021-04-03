<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="index.aspx.cs" Inherits="ImportExportTest.index" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>ImportExportTest</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <p>
                <a href="ImportExportTest.xlsx">範本下載</a>，<asp:FileUpload ID="FileUpload1" runat="server" />
                <asp:Label ID="FileUploadErr" runat="server" ForeColor="Red" Visible="False"></asp:Label>
            </p>
            <p>
                <asp:Button ID="Button1" runat="server" Text="LinqToExcel 匯入" OnClick="Button1_Click" />
            </p>
            <p>
                <asp:Button ID="Button2" runat="server" Text="EPPlus 匯入" OnClick="Button2_Click" />&nbsp;
                <asp:Button ID="Button3" runat="server" Text="EPPlus 匯出" OnClick="Button3_Click" />&nbsp;
                <asp:Button ID="Button4" runat="server" Text="EPPlus Zip 匯出" OnClick="Button4_Click" />
            </p>
            <p>
                <asp:Button ID="Button5" runat="server" Text="MiniExcel 匯入" OnClick="Button5_Click" />&nbsp;
                <asp:Button ID="Button6" runat="server" Text="MiniExcel 匯出" OnClick="Button6_Click" />
            </p>
        </div>
    </form>
</body>
</html>
