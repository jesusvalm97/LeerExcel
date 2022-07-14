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

        //const LeerExcel = () => {
        //    //Create the http request
        //    var httpRequest = new XMLHttpRequest();
        //    //Create the url
        //    var hiddenb64 = document.getElementById("HiddenB64");
        //    var url = "https://localhost:44351/LeerExcelApi.ashx?b64=" + hiddenb64.value;

        //    //Set the function for call when the state change
        //    httpRequest.onreadystatechange = function () {

        //        //If the ready state is equal to 4 and the stauts is equal to 200, so the process is successfull
        //        if (this.readyState == 4 && this.status == 200) {
        //            console.log(this);
        //        }
        //        else {
        //            console.log(this);
        //        }
        //    }

        //    //Open the and send the request
        //    httpRequest.open('GET', url, true);
        //    httpRequest.send();
        //};
        //LeerExcel();

    </script>
</body>
</html>
