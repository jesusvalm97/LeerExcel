using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LeerExcel
{
    /// <summary>
    /// Lee el contenido es un excel para retornar un json con las valores del archivo.
    /// Recibe un b64.
    /// </summary>
    public class LeerExcelApi : IHttpHandler
    {
        HttpContext Context;
        string RutaArchivoExcel;

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
            }
            catch (Exception exception)
            {
                Responder(exception.ToString());
            }
        }

        #region Methods

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

        #endregion
    }
}