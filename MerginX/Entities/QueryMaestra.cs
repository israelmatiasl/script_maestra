using System;
using FileHelpers;
using MerginX.Helpers;

namespace MerginX.Entities
{
    [DelimitedRecord("\t")]
    public class QueryMaestra
    {
        public string CodGrupoDescuento { get; set; }
        public string CodDescuento { get; set; }
        public string CodGrupoFuenteDescuento { get; set; }
        public string DesCodGrupoFuenteDescuento { get; set; }
        public string CodFuenteDescuento { get; set; }
        public string DesCodFuenteDescuento { get; set; }
        public string FlagEstado { get; set; }
        public string ImgDescuento { get; set; }
        public string ImgLogoEmpresa { get; set; }
        public string ImgLogoFuente { get; set; }
        public string TituloDescuento { get; set; }
        public string CodEmpresa { get; set; }
        public string DesCodEmpresa { get; set; }
        public string Porcentaje { get; set; }
        public string PrecioFinal { get; set; }
        public string Moneda { get; set; }
        public string DescripcionResumen { get; set; }
        public string FechaInicio { get; set; }
        public string FechaFin { get; set; }
        public string DetalleDescuento { get; set; }
        public string Restricciones { get; set; }
        public string DetalleAdicional { get; set; }
        public string DiasAtencion { get; set; }
        public string Horario { get; set; }
        public string Ruc { get; set; }
        public string RazonSocial { get; set; }
        public string Direccion { get; set; }
        public string Latitud { get; set; }
        public string Longitud { get; set; }
        public string Distrito { get; set; }
        public string Provincia { get; set; }
        public string Departamento { get; set; }
        public string FlagOnline { get; set; }
        public string SegmentoCliente { get; set; }
        public string Url { get; set; }
        public string Sexo { get; set; }
        public string TipoSubscripcion { get; set; }
        public string CodSeccion { get; set; }
        public string DesCodSeccion { get; set; }
        public string ImgLogoSeccion { get; set; }
        public string MensajeMotivador { get; set; }
        public string CodRubroParaTi { get; set; }
        public string DesCodRubroParaTi { get; set; }
        public string CodRubroParaTiModelo { get; set; }
        public string CodRubroParaTiApp { get; set; }
        public string DesCodRubroParaTiApp { get; set; }
        public string FechaProceso { get; set; }
        public string FechaEjecucion { get; set; }
        public string HoraEjecucion { get; set; }
        public string UsuarioProceso { get; set; }

        public QueryMaestra()
        {

        }

        public QueryMaestra(string codGrupoDescuento, string codDescuento, string codGrupoFuenteDescuento,
                            string desCodGrupoFuenteDescuento, string codFuenteDescuento, string desCodFuenteDescuento,
                            string flagEstado, string imgDescuento, string imgLogoEmpresa, string imgLogoFuente,
                            string tituloDescuento, string codEmpresa, string desCodEmpresa, string porcentaje,
                            string precioFinal, string moneda, string descripcionResumen, string fechaInicio,
                            string fechaFin, string detalleDescuento, string restricciones, string detalleAdicional,
                            string diasAtencion, string horario, string ruc, string razonSocial, string direccion,
                            string latitud, string longitud, string distrito, string provincia, string departamento,
                            string flagOnline, string segmentoCliente, string url, string sexo, string tipoSubscripcion,
                            string codSeccion, string desCodSeccion, string imgLogoSeccion, string mensajeMotivador,
                            string codRubroParaTi, string desCodRubroParaTi, string codRubroParaTiModelo, string codRubroParaTiApp,
                            string desCodRubroParaTiApp, string fechaProceso, string fechaEjecucion, string horaEjecucion, string usuarioProceso)
        {
            CodGrupoDescuento = codGrupoDescuento;
            CodDescuento = codDescuento;
            CodGrupoFuenteDescuento = codGrupoFuenteDescuento;
            DesCodGrupoFuenteDescuento = desCodGrupoFuenteDescuento;
            CodFuenteDescuento = codFuenteDescuento;
            DesCodFuenteDescuento = desCodFuenteDescuento;
            FlagEstado = flagEstado;
            ImgDescuento = imgDescuento;
            ImgLogoEmpresa = imgLogoEmpresa;
            ImgLogoFuente = imgLogoFuente;
            TituloDescuento = tituloDescuento;
            CodEmpresa = codEmpresa;
            DesCodEmpresa = desCodEmpresa;
            Porcentaje = porcentaje;
            PrecioFinal = precioFinal;
            Moneda = moneda;
            DescripcionResumen = descripcionResumen;
            FechaInicio = fechaInicio;
            FechaFin = Functions.ConvertToDateFromRegexDate(fechaFin);
            DetalleDescuento = detalleDescuento;
            Restricciones = restricciones;
            DetalleAdicional = detalleAdicional;
            DiasAtencion = diasAtencion;
            Horario = horario;
            Ruc = ruc;
            RazonSocial = razonSocial;
            Direccion = direccion;
            Latitud = latitud;
            Longitud = longitud;
            Distrito = distrito;
            Provincia = provincia;
            Departamento = departamento;
            FlagOnline = flagOnline;
            SegmentoCliente = segmentoCliente;
            Url = url;
            Sexo = sexo;
            TipoSubscripcion = tipoSubscripcion;
            CodSeccion = codSeccion;
            DesCodSeccion = desCodSeccion;
            ImgLogoSeccion = imgLogoSeccion;
            MensajeMotivador = mensajeMotivador;
            CodRubroParaTi = codRubroParaTi;
            DesCodRubroParaTi = desCodRubroParaTi;
            CodRubroParaTiModelo = codRubroParaTiModelo;
            CodRubroParaTiApp = codRubroParaTiApp;
            DesCodRubroParaTiApp = desCodRubroParaTiApp;
            FechaProceso = fechaProceso;
            FechaEjecucion = fechaEjecucion;
            HoraEjecucion = horaEjecucion;
            UsuarioProceso = usuarioProceso;
        }
    }
}
