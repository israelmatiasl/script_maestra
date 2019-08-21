using System;
namespace MerginX.Entities
{
    public class MaestraFuentes
    {
        public string FuenteDescuento { get; set; }
        public string ImgLogoFuente { get; set; }
        public string IdFuenteDescuento { get; set; }
        public string TipoSuscripcion { get; set; }

        public MaestraFuentes(string fuenteDescuento, string imgLogoFuente, string idFuenteDescuento, string tipoSuscripcion)
        {
            FuenteDescuento = fuenteDescuento;
            ImgLogoFuente = imgLogoFuente;
            IdFuenteDescuento = idFuenteDescuento;
            TipoSuscripcion = tipoSuscripcion;
        }
    }
}
