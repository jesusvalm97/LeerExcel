using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LeerExcel
{
    public class Class1
    {
        public Class1()
        {
            Propiedades = new Dictionary<string, object>();
        }

        /// <summary>
        /// El key es el nombre de la propiedad y el value el valor de dicha propiedad
        /// </summary>
        public Dictionary<string, object> Propiedades { get; }
    }
}