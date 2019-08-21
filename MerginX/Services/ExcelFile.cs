using System;
using System.Collections.Generic;
using FileHelpers;
using MerginX.Entities;

namespace MerginX.Services
{
    public static class ExcelFile
    {
        public static List<Prueba> readTabFile()
        {
            var engine = new FileHelperAsyncEngine<Prueba>();

            // Read
            using (engine.BeginReadFile("/Users/israelmatiasl/PRINCIPAL/Proyectos/SFTP/190725/prueba.tsv"))
            {
                // The engine is IEnumerable
                foreach (Prueba cust in engine)
                {
                    // your code here
                    Console.WriteLine(cust.Columna1);
                    Console.WriteLine(cust.Columna2);
                    Console.WriteLine(cust.Columna3);
                }
            }

            return null;
        }
    }
}
