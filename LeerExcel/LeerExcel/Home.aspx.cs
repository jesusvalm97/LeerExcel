using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
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
            //HiddenB64.Value = b64;
            string b64Encrypado = Encrypt(b64);
            string url = "https://localhost:44351/LeerExcelApi.ashx?b64=" + b64Encrypado;

            WebClient webClient = new WebClient();
            string response = webClient.DownloadString(url);
        }

        /// <summary>
        /// Encriptar un string
        /// </summary>
        /// <param name="encryptString">String a encriptar</param>
        /// <returns>El string encriptado</returns>
        public string Encrypt(string encryptString)
        {
            string EncryptionKey = "a15*/3hfjHJairtk96adfsFUIh87w340y5afdm9860-04*/2w46";
            byte[] clearBytes = Encoding.Unicode.GetBytes(encryptString);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] {
                0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76
            });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    encryptString = Convert.ToBase64String(ms.ToArray());
                }
            }
            return encryptString;
        }
    }
}