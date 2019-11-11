using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MerginX.Entities;
using MerginX.Helpers;
using MerginX.Services;

namespace MerginX
{
    class Program
    {
        static void Main(string[] args)
        {
            var fechaActual = DateTime.UtcNow;
            var fechaActualEntero = Functions.ConvertToDateString(fechaActual);
            var fechaManiana = fechaActual.AddDays(1);
            var fechaManianaEntero = Functions.ConvertToDateString(fechaManiana);

            var paths = leerRutas("/mnt/fileszpatidesa/test/maestranetcore/app/script_maestra/MerginX/path/rutas.txt", fechaActual);

            //var pathDescuentosAgotados = @"C:\\ParaTi\\data\\descuentos_agotados_20191024.csv";
            var pathDescuentosAgotados = $"{paths.FirstOrDefault(x => x.Name == "AGOTADOS").Value}";
            
            //var pathMaestraDescuentos =  @"C:\\ParaTi\\data\\Maestra_Descuentos_20191107.xlsx";
            var pathMaestraDescuentos = $"{paths.FirstOrDefault(x => x.Name == "MAESTRA").Value}";
            
            //var pathMaestraDescuentosSalida = @"C:\\ParaTi\\data\\salida";
            var pathMaestraDescuentosSalida = $"{paths.FirstOrDefault(x => x.Name == "OUT").Value}";
            
            //var pathNuevosDescuentosTSV = @"C:\\ParaTi\\data\\hive_output_20191024.tsv";
            var pathNuevosDescuentosTSV = $"{paths.FirstOrDefault(x => x.Name == "NUEVOS").Value}";
            //var pathNuevosDescuentosXLS = "/Users/israelmatiasl/PRINCIPAL/Proyectos/SFTP/CORRECCION_MAESTRA/0821/query-hive-2977.xlsx";

            modificarMaestra(pathMaestraDescuentos,
                             pathDescuentosAgotados,
                             pathNuevosDescuentosTSV,
                             pathMaestraDescuentosSalida,
                             fechaActualEntero,
                             fechaManianaEntero);
        }

        static List<ExcelPath> leerRutas(string path, DateTime fechaActual)
        {
            var paths = new List<ExcelPath>();
            string[] lines = File.ReadAllLines(path);

            foreach (string line in lines)
            {
                var dum = line.Split('|');
                var pathVal = new ExcelPath(dum[0], dum[1]);
                paths.Add(pathVal);
            }

            var pathDescuentosAgotados = paths.FirstOrDefault(x => x.Name == "AGOTADOS");
            if (pathDescuentosAgotados != null)
            {
                var today = fechaActual;
                while (true)
                {
                    var todayNumero = Functions.ConvertToDateString(today);
                    var pathFinal = $"{pathDescuentosAgotados.Value}/descuentos_agotados_{todayNumero}.csv";
                    if (!File.Exists(pathFinal))
                    {
                        today = today.AddDays(-1);
                        continue;
                    }
                    pathDescuentosAgotados.Value = pathFinal;
                    break;
                }
            }

            var pathMaestra = paths.FirstOrDefault(x => x.Name == "MAESTRA");
            if (pathMaestra != null)
            {
                var today = fechaActual;
                while (true)
                {
                    var todayNumero = Functions.ConvertToDateString(today);
                    var pathFinal = $"{pathMaestra.Value}/Maestra_Descuentos_{todayNumero}.xlsx";
                    if (!File.Exists(pathFinal))
                    {
                        today = today.AddDays(-1);
                        continue;
                    }
                    pathMaestra.Value = pathFinal;
                    break;
                }
            }

            var pathNuevos = paths.FirstOrDefault(x => x.Name == "NUEVOS");
            if (pathNuevos != null)
            {
                var today = fechaActual;
                while (true)
                {
                    var todayNumero = Functions.ConvertToDateString(today);
                    var pathFinal = $"{pathNuevos.Value}/hive_output_{todayNumero}.tsv";
                    if (!File.Exists(pathFinal))
                    {
                        today = today.AddDays(-1);
                        continue;
                    }
                    pathNuevos.Value = pathFinal;
                    break;
                }
            }

            return paths;
        }

        static void modificarMaestra(string pathMaestraExcel, string pathDescuentosVencidos, string pathDescuentosNuevos, string pathMaestraSalida, int fechaActualEntero, int fechaManianaEntero)
        {
            Console.WriteLine("[{0}] Empezando a extraer informacion de los descuentos Agotados...", DateTime.Now);
            var descuentosAgotados = ExcelService.LeerDescuentosAgotados(pathDescuentosVencidos);
            Console.WriteLine("[{0}] Se terminó de extraer informacion de los descuentos Agotados...", DateTime.Now);

            Console.WriteLine("[{0}] Empezando a extraer información de los descuentos Nuevos...", DateTime.Now);
            var descuentosQueryNuevos = ExcelService.LeerQueryMaestraTSV(pathDescuentosNuevos);
            //var descuentosQueryNuevos = ExcelService.LeerQueryMaestra(pathDescuentosNuevos);
            Console.WriteLine("[{0}] Se terminó de extraer información de los descuentos Nuevos...", DateTime.Now);

            Console.WriteLine("[{0}] Empezando a extraer información de la Maestra...", DateTime.Now);
            var maestraCompleta = ExcelService.LeerMaestraCompleta(pathMaestraExcel);
            Console.WriteLine("[{0}] Se terminó de extraer información de la Maestra...", DateTime.Now);

            var descuentosMaestra = maestraCompleta.MaestraDescuentos;
            Console.WriteLine("[{0}] La maestra tiene {1} registros", DateTime.Now, descuentosMaestra.Count);

            // Obtengo los Id's de los descuentos agotados
            var descuentosAgotadosIdGrupoBeneficios = descuentosAgotados.Select(x => x.IdGrupoBeneficio).Distinct().ToList();
            Console.WriteLine("[{0}] Se obtuvo {1} Id's de grupo de beneficio agotados", DateTime.Now, descuentosAgotadosIdGrupoBeneficios.Count);

            // Modifica el flag estado a 0 los descuentos vencidos
            Console.WriteLine("[{0}] Empezando a modificar a estado 0 los descuentos agotados...", DateTime.Now);
            modificarDescuentosAgotados(ref descuentosMaestra, descuentosAgotadosIdGrupoBeneficios);
            var descuentosAgotadosTotales = descuentosMaestra.Where(x => x.FlagEstado == "0").ToList();
            Console.WriteLine("[{0}] Se ha modificado {1} registros en total a estado 0", DateTime.Now, descuentosAgotadosTotales.Count);
            Console.WriteLine("[{0}] Se terminó de modificar a estado 0 los descuentos agotados...", DateTime.Now);

            // Obtengo los codigos grupo de descuento del query
            var codGrupoDescuentosQueryMaestra = descuentosQueryNuevos.Select(x => x.CodGrupoDescuento).Distinct();

            // Obtengo los id grupo beneficios de la maestra
            var codGrupoDescuentosMaestra = descuentosMaestra.Select(x => x.IdGrupoBeneficio).Distinct();

            // Obtengo los id grupo beneficios que no se encuentran en la maestra para ser agregados
            var nuevosDescuentosPorAgregar = codGrupoDescuentosQueryMaestra.Except(codGrupoDescuentosMaestra).ToList();
            Console.WriteLine("[{0}] Se encontraron {1} nuevos descuentos singulares por agregar", DateTime.Now, nuevosDescuentosPorAgregar.Count);

            // Agrega los nuevos descuentos a la maestra
            Console.WriteLine("[{0}] Empezando a agregar los nuevos descuentos...", DateTime.Now);
            agregarDescuentosNuevos(ref descuentosMaestra, descuentosQueryNuevos, nuevosDescuentosPorAgregar, fechaActualEntero);
            Console.WriteLine("[{0}] Se terminó de agregar los nuevos descuentos...", DateTime.Now);
            
            // Revisar los descuentos con fecha fin
            var resultadoMaestra = validarFechaFinDescuento(descuentosMaestra);
            
            //Revisar duplicidad de descuentos
            var maestraUnica = eliminarDuplicado(resultadoMaestra.DescuentosAceptados);
            
            asignarIdBeneficioANulos(ref maestraUnica);
            
            // Exporta la maestra de descuentos modificado
            maestraCompleta.MaestraDescuentos = maestraUnica;
            Console.WriteLine("[{0}] La nueva Maestra ahora tiene {1} datos", DateTime.Now, maestraCompleta.MaestraDescuentos.Count);
            Console.WriteLine("[{0}] Empezando a exportar la Maestra", DateTime.Now);
            ExcelService.ExportarMaestraModificada(maestraCompleta, pathMaestraSalida, $"Maestra_Descuentos_{fechaManianaEntero}.xlsx");
            Console.WriteLine("[{0}] Se terminó de exportar la Maestra", DateTime.Now);

            if (resultadoMaestra.DescuentosRechazados.Count > 0)
            {
                Console.WriteLine("[{0}] Empezando a exportar los descuentos rechazados", DateTime.Now);
                ExcelService.ExportarRechazados(resultadoMaestra.DescuentosRechazados, pathMaestraSalida, $"descuentos_rechazados_{fechaManianaEntero}.xlsx");
                Console.WriteLine("[{0}] Se terminó de exportar los descuentos rechazados", DateTime.Now);
            }
        }

        static void modificarDescuentosAgotados(ref List<MaestraDescuentos> maestraDescuentos, List<string> idDescuentosAgotados)
        {
            var descuentosVencidos = maestraDescuentos.Where(x => idDescuentosAgotados.Contains(x.IdGrupoBeneficio)).ToList();

            foreach (var descuento in descuentosVencidos)
            {
                descuento.FlagEstado = "0";
            }
        }
        static void agregarDescuentosNuevos(ref List<MaestraDescuentos> maestraDescuentos, List<QueryMaestra> descuentosQuery, List<string> idDescuentosNuevos, int fechaActualEntero)
        {
            foreach(var idGrupoDescuento in idDescuentosNuevos)
            {
                var descuentosNuevos = descuentosQuery.Where(x => x.CodGrupoDescuento == idGrupoDescuento &&
                                                                  string.IsNullOrEmpty(x.FechaFin)).ToList();
                foreach(var descuentoNuevo in descuentosNuevos)
                {
                    //Agregar alguna condicional en caso sea necesario
                    var descuento = new MaestraDescuentos(descuentoNuevo);
                    descuento.FechaProceso = fechaActualEntero.ToString();
                    maestraDescuentos.Add(descuento);
                }
            }
        }
        static List<MaestraDescuentos> eliminarDuplicado(List<MaestraDescuentos> maestraDescuentos)
        {
            var descuentosUnicos = new List<MaestraDescuentos>();
            Console.WriteLine("[{0}] Descuentos sin revisión de duplicidad {1}", DateTime.Now, maestraDescuentos.Count);

            var descuentos = maestraDescuentos.OrderByDescending(x=> Convert.ToInt32(x.FechaProceso)).Select(x => new
            {
                x.TituloBeneficio,
                x.DescripcionResumen,
                x.Latitud,
                x.Longitud,
                x.SegmentoCliente,
                x.Url
            }).Distinct().ToList();

            Console.WriteLine("[{0}] Descuentos con revisión de duplicidad {1}", DateTime.Now, descuentos.Count);

            foreach (var descuento in descuentos)
            {
                var descuentoUnico = maestraDescuentos.FirstOrDefault(x =>
                    x.TituloBeneficio == descuento.TituloBeneficio &&
                    x.DescripcionResumen == descuento.DescripcionResumen &&
                    x.Latitud == descuento.Latitud &&
                    x.Longitud == descuento.Longitud &&
                    x.SegmentoCliente == descuento.SegmentoCliente &&
                    x.Url == descuento.Url);
                
                descuentosUnicos.Add(descuentoUnico);
            }
            
            return descuentosUnicos;
        }
        static void asignarIdBeneficioANulos(ref List<MaestraDescuentos> maestraDescuentos)
        {
            var descuentosSinIdBeneficio = maestraDescuentos.Where(x => string.IsNullOrEmpty(x.IdBeneficio)).ToList();

            int pos = 0;
            foreach (var descuento in descuentosSinIdBeneficio)
            {
                descuento.IdBeneficio = descuento.IdGrupoBeneficio + pos.ToString().PadLeft(6, '0') + "AUTOGEN";
                pos++;
            }

            Console.WriteLine("[{0}] Se actualizaron {1} descuentos que no tenían IdBeneficios", DateTime.Now, descuentosSinIdBeneficio.Count);
        }
        static MaestraResultado validarFechaFinDescuento(List<MaestraDescuentos> maestraDescuentos)
        {
            maestraDescuentos = maestraDescuentos.Where(x=>x.FlagEstado != "0").ToList();
            
            var resultado = new MaestraResultado();
            
            var accepted = new List<MaestraDescuentos>();
            var rejected = new List<MaestraDescuentos>();
            int i = 1;
            foreach (var md in maestraDescuentos)
            {
                //Console.WriteLine("Modificando "+ i);
                if (string.IsNullOrEmpty(md.FechaFin))
                {
                    accepted.Add(md);
                }
                else
                {
                    var posNotNumeric = Functions.GetIndexNotNumeric(md.FechaFin);
                    if (posNotNumeric == -1)
                    {
                        try
                        {
                            var _ = Functions.ConvertDateFromStringNumeric(md.FechaFin);
                            accepted.Add(md);
                        }
                        catch { rejected.Add(md); }
                    }
                    else
                    {
                        try
                        {
                            var _ = Functions.ConvertFormatDateFromString(md.FechaFin);
                            accepted.Add(md);
                        }
                        catch { rejected.Add(md); }
                    }
                }
                i++;
            }

            resultado.DescuentosAceptados = accepted;
            resultado.DescuentosRechazados = rejected;
            
            return resultado;
        }
    }

}
