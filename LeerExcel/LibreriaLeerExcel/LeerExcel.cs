using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace LibreriaLeerExcel
{
    public static class LeerExcel
    {
        #region Main Methods

        public static string  Leer(byte[] bytes) {
            string rutaExcel = GuardarExcel(bytes);
            string json = ObtenerJSONDeExcel(rutaExcel);

            
            return "";
        }

        #endregion

        #region Private methods

        private static string GuardarExcel(byte[] bytes)
        {
            #region Guardar archivo .xlsx

            #region Crear carpeta data

            string rutaCarpetaData = "C:/Desarrollo/LeerExcel/LeerExcel/LibreriaLeerExcel/Data";

            //Se crea la carpeta si no existe
            if (!System.IO.Directory.Exists(rutaCarpetaData))
                System.IO.Directory.CreateDirectory(rutaCarpetaData);

            #endregion

            #region Crear archivo excel

            //Ruta del archivo excel
            string rutaArchivoExcel = $"{rutaCarpetaData}/{Guid.NewGuid()}.xlsx";
            //Crear archivo excel
            System.IO.File.WriteAllBytes(rutaArchivoExcel, bytes);

            #endregion

            #endregion

            return rutaArchivoExcel;
        }

        /// <summary>
        /// Obtener json del contenido del archivo excel
        /// </summary>
        /// <returns>String del json armado en base al contenido del archivo excel</returns>
        private static string ObtenerJSONDeExcel(string rutaArchivoExcel)
        {
            //Crear el objeto excel en base a la ruta del archivo. Es necesario crear cada objeto.
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlworkbook = xlApp.Workbooks.Open(rutaArchivoExcel);
            //El nuget de mircrosoft trabaja en base a 1 y no 0 como en un arreglo normal
            Excel._Worksheet xlWorkSheet = xlworkbook.Sheets[1];
            Excel.Range xlRange = xlWorkSheet.UsedRange;

            //Obtener numero de filas y columnas
            int numeroFilas = xlRange.Rows.Count;
            int numeroColumnas = xlRange.Columns.Count;

            //Recorrer el excel
            for (int x = 1; x <= numeroFilas; x++)
            {
                for (int y = 1; y <= numeroColumnas; y++)
                {
                    //Obtener valor de la celda
                    dynamic celda = xlRange.Cells[x, y];
                    dynamic valorCelda = celda.Value2;
                }
            }

            return "";
        }

        #endregion
    }
}
