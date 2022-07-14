using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace LeerExcel
{
    public partial class Home : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(@"C:\Users\e4911449\Documents\EjemploExcel.xlsx");
            string b64 = Convert.ToBase64String(bytes);
            HiddenB64.Value = b64;
            string url = "https://localhost:44351/LeerExcelApi.ashx?b64=" + b64;

            WebClient webClient = new WebClient();
            string response = webClient.DownloadString(url);
        }
    }
}