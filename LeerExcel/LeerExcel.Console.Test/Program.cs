using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeerExcel.Console.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(@"C:\Users\e4911449\Documents\EjemploExcel.xlsx");
            LibreriaLeerExcel.LeerExcel.Leer(bytes);
        }
    }
}
