using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ApiLeerExcel.Models
{
    public class ObjetoRespuesta
    {
        public ObjetoRespuesta()
        {
            Propiedades = new Dictionary<string, object>();
        }

        /// <summary>
        /// El key es el nombre de la propiedad y el value el valor de dicha propiedad
        /// </summary>
        public Dictionary<string, object> Propiedades { get; }
    }
}
