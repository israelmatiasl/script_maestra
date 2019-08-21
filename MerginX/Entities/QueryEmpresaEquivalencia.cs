using System;
namespace MerginX.Entities
{
    public class QueryEmpresaEquivalencia
    {
        public string DesCodFuenteDescuento { get; set; }
        public string DesCodFuenteEmpresa { get; set; }
        public string DireccionFuente { get; set; }
        public string DireccionApi { get; set; }
        public string Direccion { get; set; }
        public string Latitud { get; set; }
        public string Longitud { get; set; }
        public string Distrito { get; set; }
        public string Provincia { get; set; }
        public string Departamento { get; set; }
        public string TipoNuevo { get; set; }
        public string FechaEjecucion { get; set; }
        public string HoraEjecucion { get; set; }
        public string UsuarioProceso { get; set; }
        public string FechaProceso { get; set; }

        public QueryEmpresaEquivalencia(string desCodFuenteDescuento, string desCodFuenteEmpresa, string direccionFuente,
                                        string direccionApi, string direccion, string latitud, string longitud, string distrito,
                                        string provincia, string departamento, string tipoNuevo, string fechaEjecucion,
                                        string horaEjecucion, string usuarioProceso, string fechaProceso)
        {
            DesCodFuenteDescuento   = desCodFuenteDescuento;
            DesCodFuenteEmpresa     = desCodFuenteEmpresa;
            DireccionFuente         = direccionFuente;
            DireccionApi            = direccionApi;
            Direccion               = direccion;
            Latitud                 = latitud;
            Longitud                = longitud;
            Distrito                = distrito;
            Provincia               = provincia;
            Departamento            = departamento;
            TipoNuevo               = tipoNuevo;
            FechaEjecucion          = fechaEjecucion;
            HoraEjecucion           = horaEjecucion;
            UsuarioProceso          = usuarioProceso;
            FechaProceso            = fechaProceso;
        }
    }
}
