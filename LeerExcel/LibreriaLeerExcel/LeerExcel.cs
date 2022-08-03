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

        /// <summary>
        /// Leer un excel medio de los bytes
        /// </summary>
        /// <param name="bytes">Los bytes del excel a leer</param>
        /// <returns>Un json con el contenido del excel. El json está armado en base a que cada fila es un objeto y cada columna es una propiedad</returns>
        public static string Leer(byte[] bytes)
        {
            string rutaExcel = GuardarExcel(bytes);
            string json = ObtenerJSONDeExcel(rutaExcel);

            return json;
        }

        /// <summary>
        /// Leer un excel medio del base 64
        /// </summary>
        /// <param name="b64">El base 64 del excel a leer</param>
        /// <returns>Un json con el contenido del excel. El json está armado en base a que cada fila es un objeto y cada columna es una propiedad</returns>
        public static string Leer(string b64)
        {
            //Convertir base 64 a arreglo de bytes
            byte[] bytes = Convert.FromBase64String(b64);

            return Leer(bytes);
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

            //Diccionario para mantener el control de las propiedades de cada objeto/fila
            //Key es el numero de la columna y el valor es el nombre de la columna/propiedad
            Dictionary<int, string> propiedades = new Dictionary<int, string>();

            //Lista de objetos con los valores del excel ya armado por filas. Cada objeto es una fila, cada propiedad es una columna
            List<Objeto> objetos = new List<Objeto>();

            //Recorrer el excel
            for (int x = 1; x <= numeroFilas; x++)
            {
                //Armando objeto con los valores de la fila
                Objeto objeto = new Objeto();

                for (int y = 1; y <= numeroColumnas; y++)
                {
                    //Obtener valor de la celda
                    dynamic celda = xlRange.Cells[x, y];
                    dynamic valorCelda = celda.Value2;

                    //Si x es igual a 1, significa que es la fila 1 y es la que corresponde al nombre de las propiedades
                    if (x == 1)
                    {
                        propiedades.Add(y, valorCelda);
                        continue;
                    }

                    //Obtener nombre de la propiedad/columna
                    string nombrePropiedad = propiedades.Where(m => m.Key == y).FirstOrDefault().Value;
                    objeto.Propiedades.Add(nombrePropiedad, valorCelda);
                }

                //Agregar a la lista el objeto solo cuando x sea mayor a 1, porque la fila 1 es la de el nombre de las columnas
                if (x > 1)
                    objetos.Add(objeto);
            }

            //Retornar la lista de objetos en forma de json
            return Newtonsoft.Json.JsonConvert.SerializeObject(objetos);
        }

        #endregion
    }
}
