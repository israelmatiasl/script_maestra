using System;
using MerginX.Helpers;

namespace MerginX.Entities
{
    public class MaestraDescuentos
    {
        public string IdGrupoBeneficio { get; set; }
        public string IdBeneficio { get; set; }
        public string FuenteDescuentoGrupo { get; set; }
        public string IdFuenteDescuento { get; set; }
        public string FuenteDescuento { get; set; }
        public string FlagEstado { get; set; }
        public string ImgDescuento { get; set; }
        public string ImgLogoEmpresa { get; set; }
        public string ImgLogoFuente { get; set; }
        public string TituloBeneficio { get; set; }
        public string IdEstablecimiento { get; set; }
        public string NombreEstablecimiento { get; set; }
        public string Porcentaje { get; set; }
        public string PrecioFinal { get; set; }
        public string Moneda { get; set; }
        public string DescripcionResumen { get; set; }
        public string FechaInicio { get; set; }
        public string FechaFin { get; set; }
        public string DetalleBeneficio { get; set; }
        public string Restricciones { get; set; }
        public string DetalleAdicional { get; set; }
        public string DiasAtencionEstab { get; set; }
        public string HorarioEstablec { get; set; }
        public string Ruc { get; set; }
        public string RazonSocial { get; set; }
        public string Direccion { get; set; }
        public string Latitud { get; set; }
        public string Longitud { get; set; }
        public string FlagOnline { get; set; }
        public string RubroBcpApp { get; set; }
        public string RubroParaTi { get; set; }
        public string RubroParaTiApp { get; set; }
        public string DepartamentoEstablec { get; set; }
        public string ProvinciaEstablec { get; set; }
        public string DistritoEstablec { get; set; }
        public string SegmentoCliente { get; set; }
        public string Url { get; set; }
        public string Sexo { get; set; }
        public string MensajeMotivador { get; set; }
        public string TipoSubscripcion { get; set; }
        public string TipoOrigen { get; set; }
        public string CodPromocion { get; set; }
        public string FechaProceso { get; set; }

        public MaestraDescuentos()
        {

        }

        public MaestraDescuentos(string idGrupoBeneficio, string idBeneficio, string fuenteDescuentoGrupo,
                                 string idFuenteDescuento, string fuenteDescuento, string flagEstado, string imgDescuento,
                                 string imgLogoEmpresa, string imgLogoFuente, string tituloBeneficio, string idEstablecimiento,
                                 string nombreEstablecimiento, string porcentaje, string precioFinal, string moneda,
                                 string descripcionResumen, string fechaInicio, string fechaFin, string detalleBeneficio,
                                 string restricciones, string detalleAdicional, string diasAtencionEstab,
                                 string horarioEstablec, string ruc, string razonSocial, string direccion, string latitud,
                                 string longitud, string flagOnline, string rubroBcpApp, string rubroParaTi, string rubroParaTiApp,
                                 string departamentoEstablec, string provinciaEstablec, string distritoEstablec,
                                 string segmentoCliente, string url, string sexo, string mensajeMotivador,
                                 string tipoSubscripcion, string tipoOrigen, string codPromocion, string fechaProceso)
        {
            IdGrupoBeneficio        = idGrupoBeneficio;
            IdBeneficio             = idBeneficio;
            FuenteDescuentoGrupo    = fuenteDescuentoGrupo;
            IdFuenteDescuento       = idFuenteDescuento;
            FuenteDescuento         = fuenteDescuento;
            FlagEstado              = flagEstado;
            ImgDescuento            = imgDescuento;
            ImgLogoEmpresa          = imgLogoEmpresa;
            ImgLogoFuente           = imgLogoFuente;
            TituloBeneficio         = tituloBeneficio;
            IdEstablecimiento       = idEstablecimiento;
            NombreEstablecimiento   = nombreEstablecimiento;
            Porcentaje              = porcentaje;
            PrecioFinal             = precioFinal;
            Moneda                  = moneda;
            DescripcionResumen      = descripcionResumen;
            FechaInicio             = fechaInicio;
            FechaFin                = fechaFin;
            DetalleBeneficio        = detalleBeneficio;
            Restricciones           = restricciones;
            DetalleAdicional        = detalleAdicional;
            DiasAtencionEstab       = diasAtencionEstab;
            HorarioEstablec         = horarioEstablec;
            Ruc                     = ruc;
            RazonSocial             = razonSocial;
            Direccion               = direccion;
            Latitud                 = latitud;
            Longitud                = longitud;
            FlagOnline              = flagOnline;
            RubroBcpApp             = rubroBcpApp;
            RubroParaTi             = rubroParaTi;
            RubroParaTiApp          = rubroParaTiApp;
            DepartamentoEstablec    = departamentoEstablec;
            ProvinciaEstablec       = provinciaEstablec;
            DistritoEstablec        = distritoEstablec;
            SegmentoCliente         = segmentoCliente?.ToUpper();
            Url                     = url;
            Sexo                    = sexo;
            MensajeMotivador        = mensajeMotivador;
            TipoSubscripcion        = tipoSubscripcion;
            TipoOrigen              = tipoOrigen;
            CodPromocion            = codPromocion;
            FechaProceso            = fechaProceso;
        }

        public MaestraDescuentos(QueryMaestra queryMaestra)
        {
            IdGrupoBeneficio        = queryMaestra.CodGrupoDescuento;
            IdBeneficio             = queryMaestra.CodDescuento;
            FuenteDescuentoGrupo    = queryMaestra.DesCodGrupoFuenteDescuento;
            IdFuenteDescuento       = queryMaestra.CodFuenteDescuento;
            FuenteDescuento         = queryMaestra.DesCodFuenteDescuento;
            FlagEstado              = queryMaestra.FlagEstado;
            ImgDescuento            = queryMaestra.ImgDescuento;
            ImgLogoEmpresa          = queryMaestra.ImgLogoEmpresa;
            ImgLogoFuente           = queryMaestra.ImgLogoFuente;
            TituloBeneficio         = queryMaestra.TituloDescuento;
            IdEstablecimiento       = queryMaestra.CodEmpresa;
            NombreEstablecimiento   = queryMaestra.DesCodEmpresa;
            Porcentaje              = queryMaestra.Porcentaje;
            PrecioFinal             = queryMaestra.PrecioFinal;
            Moneda                  = queryMaestra.Moneda;
            DescripcionResumen      = queryMaestra.DescripcionResumen;
            FechaInicio             = queryMaestra.FechaInicio;
            FechaFin                = queryMaestra.FechaFin;
            DetalleBeneficio        = queryMaestra.DetalleDescuento;
            Restricciones           = queryMaestra.Restricciones;
            DetalleAdicional        = queryMaestra.DetalleAdicional;
            DiasAtencionEstab       = queryMaestra.DiasAtencion;
            HorarioEstablec         = queryMaestra.Horario;
            Ruc                     = queryMaestra.Ruc;
            RazonSocial             = queryMaestra.RazonSocial;
            Direccion               = queryMaestra.Direccion;
            Latitud                 = queryMaestra.Latitud;
            Longitud                = queryMaestra.Longitud;
            FlagOnline              = queryMaestra.FlagOnline;
            RubroBcpApp             = queryMaestra.DesCodSeccion;
            RubroParaTi             = queryMaestra.DesCodRubroParaTi;
            RubroParaTiApp          = queryMaestra.DesCodRubroParaTiApp;
            DepartamentoEstablec    = queryMaestra.Departamento;
            ProvinciaEstablec       = queryMaestra.Provincia;
            DistritoEstablec        = queryMaestra.Distrito;
            SegmentoCliente         = queryMaestra.SegmentoCliente != null ? queryMaestra.SegmentoCliente.ToUpper() : queryMaestra.SegmentoCliente;
            Url                     = queryMaestra.Url;
            Sexo                    = queryMaestra.Sexo;
            MensajeMotivador        = queryMaestra.MensajeMotivador;
            TipoSubscripcion        = queryMaestra.TipoSubscripcion;
            TipoOrigen              = string.Empty;
            CodPromocion            = string.Empty;
            FechaProceso            = queryMaestra.FechaProceso;
        }
    }
}
