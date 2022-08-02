using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;
using System.IO;

namespace LeerExcel
{
    /// <summary>
    /// Lee el contenido es un excel para retornar un json con las valores del archivo.
    /// Recibe un b64.
    /// </summary>
    public class LeerExcelApi : IHttpHandler
    {
        #region Propiedades

        /// <summary>
        /// Context sirve para obtener los parametros del request y para retornar el response
        /// </summary>
        HttpContext Context
        {
            get;
            set;
        }

        /// <summary>
        /// La ruta temporal del archivo excel
        /// </summary>
        string RutaArchivoExcel
        {
            get
            {
                object value = Context.Session["LeerExcelApi.RutaArchivoExcel"];

                if (value == null)
                {
                    return string.Empty;
                }

                return value.ToString();
            }
            set
            {
                Context.Session["LeerExcelApi.RutaArchivoExcel"] = value;
            }
        }

        /// <summary>
        /// Llave para desencriptar
        /// </summary>
        //string EncryptionKey
        //{
        //    get
        //    {
        //        object value = Context.Session["LeerExcelApi.EncryptionKey"];

        //        if (value == null)
        //        {
        //            return string.Empty;
        //        }

        //        return value.ToString();
        //    }
        //    set
        //    {
        //        Context.Session["LeerExcelApi.EncryptionKey"] = value;
        //    }
        //}

        #endregion

        public void ProcessRequest(HttpContext context)
        {
            Context = context;

            try
            {
                #region Obtener parametros

                //Obtener el b64
                string b64Excel = context.Request.QueryString["b64"];

                //Validar de que el b64 no venga vacio
                if (string.IsNullOrEmpty(b64Excel))
                {
                    Responder("El argumento b64 es nulo.");
                }

                //Guardar archivo excel para manipularlo
                if (!GuardarExcel(b64Excel))
                {
                    Responder("No se pudo guardar el archivo excel.");
                }

                #endregion

                //Obtener json del contenido del excel
                string json = ObtenerJSONDeExcel();

                Responder(json);
            }
            catch (Exception exception)
            {
                Responder(exception.ToString());
            }
        }

        #region Metodos

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// Responder al cliente
        /// </summary>
        /// <param name="respuesta">Respuesta al cliente</param>
        public void Responder(string respuesta)
        {
            Context.Response.ClearContent();
            Context.Response.Clear();
            Context.Response.Write(respuesta);
        }

        /// <summary>
        /// Obtiene los bytes del excel en base al base 64. Crea temporalmente el archivo en ./Data/NombreArchivo.xlsx para su uso, para luego eliminarlo
        /// </summary>
        /// <param name="b64">El base 64 del archivo excel.</param>
        public bool GuardarExcel(string b64)
        {
            //Validar que el b64 no esté vacío
            if (string.IsNullOrEmpty(b64))
            {
                return false;
            }

            b64 = Decrypt(b64);

            //Convertir b64 a bytes
            byte[] bytes = Convert.FromBase64String(b64);

            #region Guardar archivo .xlsx

            #region Crear carpeta data

            string rutaCarpetaData = Context.Server.MapPath("Data");

            //Se crea la carpeta si no existe
            if (!System.IO.Directory.Exists(rutaCarpetaData))
                System.IO.Directory.CreateDirectory(rutaCarpetaData);

            #endregion

            #region Crear archivo excel

            //Ruta del archivo excle
            RutaArchivoExcel = $"{rutaCarpetaData}/{Guid.NewGuid()}.xlsx";
            //Crear archivo excel
            System.IO.File.WriteAllBytes(RutaArchivoExcel, bytes);

            #endregion

            #endregion

            return true;
        }

        /// <summary>
        /// Obtener json del contenido del archivo excel
        /// </summary>
        /// <returns>String del json armado en base al contenido del archivo excel</returns>
        public string ObtenerJSONDeExcel()
        {
            //Crear el objeto excel en base a la ruta del archivo. Es necesario crear cada objeto.
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlworkbook = xlApp.Workbooks.Open(RutaArchivoExcel);
            //El nuget de mircrosoft trabaja en base a 1 y no 0 como en un arreglo normal
            Excel._Worksheet xlWorkSheet = xlworkbook.Sheets[1];
            Excel.Range xlRange = xlWorkSheet.UsedRange;

            //Obtener numero de filas y columnas
            int numeroFilas = xlRange.Rows.Count;
            int numeroColumnas = xlRange.Columns.Count;

            //Recorrer el excel
            for (int x = 1; x < numeroFilas; x++)
            {
                for (int y = 1; y < numeroColumnas; y++)
                {
                    //Obtener valor de la celda
                    dynamic celda = xlRange.Cells[x, y];
                    dynamic valorCelda = celda.Value2;
                }
            }

            return "";
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

        /// <summary>
        /// Desencripta el string
        /// </summary>
        /// <param name="cipherText">String encriptado</param>
        /// <returns>El string desencriptado</returns>
        public string Decrypt(string cipherText)
        {
            string EncryptionKey = "a15*/3hfjHJairtk96adfsFUIh87w340y5afdm9860-04*/2w46";
            cipherText = cipherText.Replace(" ", "+");
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] {
                0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76
            });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }


        #endregion
    }
}