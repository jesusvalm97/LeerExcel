using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibreriaLeerExcel
{
    public class Objeto
    {
        public Objeto()
        {
            Propiedades = new Dictionary<string, object>();
        }

        public Dictionary<string, object> Propiedades { get; }
    }
}
