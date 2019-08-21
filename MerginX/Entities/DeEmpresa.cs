using System;
namespace MerginX.Entities
{
    public class DeEmpresa
    {
        public string NumeroFila { get; set; }
        public string CodEmpresa { get; set; }
        public string CodLocal { get; set; }
        public string DesCodEmpresa { get; set; }
        public string DireccionFuente { get; set; }
        public string DireccionApi { get; set; }
        public string Direccion { get; set; }
        public string Latitud { get; set; }
        public string Longitud { get; set; }
        public string Distrito { get; set; }
        public string Provincia { get; set; }
        public string Departamento { get; set; }
        public string FechaEjecucion { get; set; }
        public string HoraEjecucion { get; set; }
        public string UsuarioProceso { get; set; }
        public string FechaProceso { get; set; }

        public DeEmpresa(string numeroFila, string codEmpresa, string codLocal, string desCodEmpresa, string direccionFuente, string direccionApi,
                         string direccion, string latitud, string longitud, string distrito, string provincia, string departamento,
                         string fechaEjecucion, string horaEjecucion, string usuarioProceso, string fechaProceso)
        {
            NumeroFila      = numeroFila;
            CodEmpresa      = codEmpresa;
            CodLocal        = codLocal;
            DesCodEmpresa   = desCodEmpresa;
            DireccionFuente = direccionFuente;
            DireccionApi    = direccionApi;
            Direccion       = direccion;
            Latitud         = latitud;
            Longitud        = longitud;
            Distrito        = distrito;
            Provincia       = provincia;
            Departamento    = departamento;
            FechaEjecucion  = fechaEjecucion;
            HoraEjecucion   = horaEjecucion;
            UsuarioProceso  = usuarioProceso;
            FechaProceso    = fechaProceso;
        }

        public DeEmpresa(QueryEmpresaEquivalencia queryEmpresaEquivalencia)
        {
            CodEmpresa      = "";
            CodLocal        = "";
            DesCodEmpresa   = queryEmpresaEquivalencia.DesCodFuenteEmpresa;
            DireccionFuente = queryEmpresaEquivalencia.DireccionFuente.Replace(",", ".");
            DireccionApi    = queryEmpresaEquivalencia.DireccionApi.Replace(",", ".");
            Direccion       = queryEmpresaEquivalencia.Direccion.Replace(",", ".");
            Latitud         = queryEmpresaEquivalencia.Latitud.Replace(",", ".");
            Longitud        = queryEmpresaEquivalencia.Longitud.Replace(",", ".");
            Distrito        = queryEmpresaEquivalencia.Distrito.Replace(",", ".");
            Provincia       = queryEmpresaEquivalencia.Provincia.Replace(",", ".");
            Departamento    = queryEmpresaEquivalencia.Departamento.Replace(",", ".");
            FechaEjecucion  = queryEmpresaEquivalencia.FechaEjecucion.Replace(",", ".");
            HoraEjecucion   = queryEmpresaEquivalencia.HoraEjecucion.Replace(",", ".");
            UsuarioProceso  = queryEmpresaEquivalencia.UsuarioProceso.Replace(",", ".");
            FechaProceso    = queryEmpresaEquivalencia.FechaProceso.Replace(",", ".");
        }
    }
}