using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApiLeerExcel.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ApiLeerExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LeerExcelT1 : ControllerBase
    {
        #region Propiedades

        /// <summary>
        /// La ruta temporal del archivo excel
        /// </summary>
        string RutaArchivoExcel
        {
            get
            {
                object value = HttpContext.Session.GetString("LeerExcelApi.RutaArchivoExcel");

                if (value == null)
                {
                    return string.Empty;
                }

                return value.ToString();
            }
            set
            {
                HttpContext.Session.SetString("LeerExcelApi.RutaArchivoExcel", value);
            }
        }

        #endregion

        [HttpPost]
        //public async Task<ActionResult<ObjetoRespuesta>> PostExcel(string b64Excel)
        //public ActionResult<ObjetoRespuesta> PostExcel(ObjetoRespuesta objetoRespuesta_Cliente)
        public ActionResult<ObjetoRespuesta> PostExcel(string b64Excel)
        {
            //Obtener b64
            //string b64Excel = objetoRespuesta_Cliente.Propiedades.Where(m => m.Key == "b64").FirstOrDefault().Value.ToString();

            //Guardar archivo excel
            //GuardarExcel(b64Excel);

            //Convertir b64 a bytes
            byte[] bytes = Convert.FromBase64String(b64Excel);

            #region Guardar archivo .xlsx

            #region Crear carpeta data

            string rutaCarpetaData = @"C:\Desarrollo\LeerExcel\LeerExcel\ApiLeerExcel\Data";

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

            //Obtener json del contenido del excel
            //string json = ObtenerJSONDeExcel();

            //Crear el objeto excel en base a la ruta del archivo. Es necesario crear cada objeto.
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlworkbook = xlApp.Workbooks.Open(RutaArchivoExcel);
            //El nuget de mircrosoft trabaja en base a 1 y no 0 como en un arreglo normal
            Excel._Worksheet xlWorkSheet = (Excel._Worksheet)xlworkbook.Sheets[1];
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

            string json = string.Empty;

            ObjetoRespuesta objetoRespuesta = new ObjetoRespuesta();
            objetoRespuesta.Propiedades.Add("Saludo", "Hola mundo");

            return Ok(objetoRespuesta);
        }

        #region Custom methods

        /// <summary>
        /// Obtiene los bytes del excel en base al base 64. Crea temporalmente el archivo en ./Data/NombreArchivo.xlsx para su uso, para luego eliminarlo
        /// </summary>
        /// <param name="b64">El base 64 del archivo excel.</param>
        //public bool GuardarExcel(string b64)
        //{
        //    //Validar que el b64 no esté vacío
        //    if (string.IsNullOrEmpty(b64))
        //    {
        //        return false;
        //    }

        //    //b64 = Decrypt(b64);

        //    //Convertir b64 a bytes
        //    byte[] bytes = Convert.FromBase64String(b64);

        //    #region Guardar archivo .xlsx

        //    #region Crear carpeta data

        //    string rutaCarpetaData = @"C:\Desarrollo\LeerExcel\LeerExcel\ApiLeerExcel\Data";

        //    //Se crea la carpeta si no existe
        //    if (!System.IO.Directory.Exists(rutaCarpetaData))
        //        System.IO.Directory.CreateDirectory(rutaCarpetaData);

        //    #endregion

        //    #region Crear archivo excel

        //    //Ruta del archivo excle
        //    RutaArchivoExcel = $"{rutaCarpetaData}/{Guid.NewGuid()}.xlsx";
        //    //Crear archivo excel
        //    System.IO.File.WriteAllBytes(RutaArchivoExcel, bytes);

        //    #endregion

        //    #endregion

        //    return true;
        //}

        /// <summary>
        /// Obtener json del contenido del archivo excel
        /// </summary>
        /// <returns>String del json armado en base al contenido del archivo excel</returns>
        //public string ObtenerJSONDeExcel()
        //{
        //    //Crear el objeto excel en base a la ruta del archivo. Es necesario crear cada objeto.
        //    Excel.Application xlApp = new Excel.Application();
        //    Excel.Workbook xlworkbook = xlApp.Workbooks.Open(RutaArchivoExcel);
        //    //El nuget de mircrosoft trabaja en base a 1 y no 0 como en un arreglo normal
        //    Excel._Worksheet xlWorkSheet = (Excel._Worksheet)xlworkbook.Sheets[1];
        //    Excel.Range xlRange = xlWorkSheet.UsedRange;

        //    //Obtener numero de filas y columnas
        //    int numeroFilas = xlRange.Rows.Count;
        //    int numeroColumnas = xlRange.Columns.Count;

        //    //Recorrer el excel
        //    for (int x = 1; x < numeroFilas; x++)
        //    {
        //        for (int y = 1; y < numeroColumnas; y++)
        //        {
        //            //Obtener valor de la celda
        //            dynamic celda = xlRange.Cells[x, y];
        //            dynamic valorCelda = celda.Value2;
        //        }
        //    }

        //    return "";
        //}

        #endregion
    }
}
