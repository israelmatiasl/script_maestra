using System;
using FileHelpers;

namespace MerginX.Entities
{
    [DelimitedRecord("\t")]
    public class Prueba
    {
        public string Columna1 { get; set; }
        public string Columna2 { get; set; }
        public string Columna3 { get; set; }

        public Prueba()
        {

        }

        public Prueba(string columna1, string columna2, string columna3)
        {
            Columna1 = columna1;
            Columna2 = columna2;
            Columna3 = columna3;
        }
    }
}
