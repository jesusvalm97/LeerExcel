<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Home.aspx.cs" Inherits="LeerExcel.Home" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            Hola mundo
            
            <asp:HiddenField ID="HiddenB64" runat="server" ClientIDMode="Static" />
            
        </div>
    </form>

    <script>
        const LeerExcel = async () => {
            try {

                var hiddenb64 = document.getElementById("HiddenB64");

                const resp = await fetch("/LeerExcelApi.ashx?b64=" + hiddenb64.value);
                const data = await.resp.json();
                console.log(data);
            }
            catch (error) {
                console.log(error);
            }
        };
    </script>
</body>
</html>
