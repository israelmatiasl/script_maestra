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
            //var pathDeEmpresa = "/Users/israelmatiasl/PRINCIPAL/Proyectos/SFTP/CORRECCION_MAESTRA/de_empresa_excel.xlsx";
            // De Empresa Leer
            //leerDeEmpresa(pathDeEmpresa);
            var fechaActual = DateTime.UtcNow;
            var fechaActualEntero = Functions.ConvertToDateString(fechaActual);
            var fechaManiana = fechaActual.AddDays(1);
            var fechaManianaEntero = Functions.ConvertToDateString(fechaManiana);

            var paths = leerRutas("/mnt/fileszpatidesa/test/maestranetcore/app/script_maestra/MerginX/path/rutas.txt", fechaActual);

            //var pathDescuentosAgotados = $"/Users/israelmatiasl/PRINCIPAL/Proyectos/SFTP/CORRECCION_MAESTRA/0821/descuentos_agotados_{fechaActualEntero}.csv"; //$"/mnt/fileszpatidesa/test/maestranetcore/20190821/descuentos_agotados_{fechaActualEntero}.csv";// $"{paths.FirstOrDefault(x=>x.Name == "AGOTADOS").Value}/descuentos_agotados_{fechaActualEntero}.csv";
            var pathDescuentosAgotados = $"{paths.FirstOrDefault(x => x.Name == "AGOTADOS").Value}";
            //var pathMaestraDescuentos = $"/Users/israelmatiasl/PRINCIPAL/Proyectos/SFTP/CORRECCION_MAESTRA/0821/Maestra_Descuentos_20190724.xlsx"; //$"/mnt/fileszpatidesa/test/maestranetcore/20190821/Maestra_Descuentos_{fechaActualEntero}.xlsx";// $"{paths.FirstOrDefault(x => x.Name == "MAESTRA").Value}/Maestra_Descuentos_{fechaActualEntero}.xlsx"; //$"/Users/israelmatiasl/PRINCIPAL/Proyectos/SFTP/CORRECCION_MAESTRA/0724/Maestra_Descuentos_{fechaActualEntero}.xlsx";
            var pathMaestraDescuentos = $"{paths.FirstOrDefault(x => x.Name == "MAESTRA").Value}";
            //var pathMaestraDescuentosSalida = $"/Users/israelmatiasl/PRINCIPAL/Proyectos/SFTP/CORRECCION_MAESTRA/0821/salida"; //$"/mnt/fileszpatidesa/test/maestranetcore/20190821";// $"{paths.FirstOrDefault(x => x.Name == "OUT").Value}";//"/Users/israelmatiasl/PRINCIPAL/Proyectos/SFTP/CORRECCION_MAESTRA/0724/salida";
            var pathMaestraDescuentosSalida = $"{paths.FirstOrDefault(x => x.Name == "OUT").Value}";
            //var pathNuevosDescuentosTSV = $"/Users/israelmatiasl/PRINCIPAL/Proyectos/SFTP/CORRECCION_MAESTRA/20190821/hive_output_{fechaActualEntero}.tsv"; //$"/mnt/fileszpatidesa/test/maestranetcore/20190821/hive_output_{fechaActualEntero}.tsv";// $"{paths.FirstOrDefault(x => x.Name == "NUEVOS").Value}/hive_output_{fechaActualEntero}.tsv";// $"/Users/israelmatiasl/PRINCIPAL/Proyectos/SFTP/190723/hive_output.tsv";
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

            // Exporta la maestra de descuentos modificado
            Console.WriteLine("..{0}", descuentosMaestra.Count);
            maestraCompleta.MaestraDescuentos = descuentosMaestra.Where(x=>x.FlagEstado != "0").ToList();
            Console.WriteLine("[{0}] La nueva Maestra ahora tiene {1} datos", DateTime.Now, maestraCompleta.MaestraDescuentos.Count);
            Console.WriteLine("[{0}] Empezando a exportar la Maestra", DateTime.Now);
            ExcelService.ExportarMaestraModificada(maestraCompleta, pathMaestraSalida, fechaManianaEntero);
            Console.WriteLine("[{0}] Se terminó de exportar la Maestra", DateTime.Now);
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
                                                                    ((string.IsNullOrEmpty(x.FechaFin)) ||
                                                                    (!string.IsNullOrEmpty(x.FechaFin) && Convert.ToInt32(x.FechaFin) >= fechaActualEntero))
                                                             ).ToList();

                foreach(var descuentoNuevo in descuentosNuevos)
                {
                    //Agregar alguna condicional en caso sea necesario
                    var descuento = new MaestraDescuentos(descuentoNuevo);
                    descuento.FechaProceso = fechaActualEntero.ToString();
                    maestraDescuentos.Add(descuento);
                }
            }
        }

        static void leerDeEmpresa(string pathDeEmpresa)
        {
            var listDeEmpresa = ExcelService.LeerDeEmpresaExcel(pathDeEmpresa);

            var listDeEmpresaGroup = listDeEmpresa.GroupBy(x => new { x.DesCodEmpresa, x.DireccionFuente }).ToList();

            foreach (var grouping in listDeEmpresaGroup)
            {
                if (grouping.Count() > 1)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    
                    Console.WriteLine("DesCodEmpresa: {0}", grouping.First().DesCodEmpresa);
                    Console.WriteLine("DireccionFuente: {0}", grouping.First().DireccionFuente);
                    foreach (var group in grouping)
                    {
                        Console.ForegroundColor = ConsoleColor.Blue;
                        Console.WriteLine("({0}) - Dirección: {1}", group.NumeroFila, group.Direccion);
                    }
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("---------------------------------------------------------------------------------------------------------------------------------");
                }
            }
        }
    }

}
