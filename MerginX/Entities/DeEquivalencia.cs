using System;
namespace MerginX.Entities
{
    public class DeEquivalencia
    {
        public string CodEmpresa { get; set; }
        public string DesCodEmpresa { get; set; }
        public string ImgLogoEmpresa { get; set; }
        public string DesCodEmpresaFuente { get; set; }
        public string FechaEjecucion { get; set; }
        public string HoraEjecucion { get; set; }
        public string UsuarioProceso { get; set; }
        public string FechaProceso { get; set; }

        public DeEquivalencia(string codEmpresa, string desCodEmpresa, string imgLogoEmpresa, string desCodEmpresaFuente,
                              string fechaEjecucion, string horaEjecucion, string usuarioProceso, string fechaProceso)
        {
            CodEmpresa = codEmpresa;
            DesCodEmpresa = desCodEmpresa;
            ImgLogoEmpresa = imgLogoEmpresa;
            DesCodEmpresaFuente = desCodEmpresaFuente;
            FechaEjecucion = fechaEjecucion;
            HoraEjecucion = horaEjecucion;
            UsuarioProceso = usuarioProceso;
            FechaProceso = fechaProceso;
        }
    }
}
