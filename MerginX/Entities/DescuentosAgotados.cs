using System;
namespace MerginX.Entities
{
    public class DescuentosAgotados
    {
        public int? FlagNuevo { get; set; }
        public string IdGrupoBeneficio { get; set; }
        public string Url { get; set; }

        public DescuentosAgotados(int? flagNuevo, string idGrupoBeneficio, string url)
        {
            FlagNuevo           = flagNuevo;
            IdGrupoBeneficio    = idGrupoBeneficio;
            Url                 = url;
        }
    }
}
