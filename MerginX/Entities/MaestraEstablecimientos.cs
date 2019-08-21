using System;
namespace MerginX.Entities
{
    public class MaestraEstablecimientos
    {
        public string IdEstablecimiento { get; set; }
        public string NombreEstablecimiento { get; set; }
        public string Formato { get; set; }
        public string ImgLogoEmpresa { get; set; }
        public string Observacion { get; set; }
        public string ImgLogoEmpresaOld { get; set; }


        public MaestraEstablecimientos(string idEstablecimiento, string nombreEstablecimiento, string formato, string imgLogoEmpresa, string observacion, string imgLogoEmpresaOld)
        {
            IdEstablecimiento = idEstablecimiento;
            NombreEstablecimiento = nombreEstablecimiento;
            Formato = formato;
            ImgLogoEmpresa = imgLogoEmpresa;
            Observacion = observacion;
            ImgLogoEmpresaOld = imgLogoEmpresaOld;
        }
    }
}
