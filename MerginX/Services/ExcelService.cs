using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using FileHelpers;
using MerginX.Entities;
using MerginX.Helpers;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MerginX.Services
{
    public static class ExcelService
    {
        static XSSFWorkbook hssfworkbook { get; set; }

        public static List<DescuentosAgotados> LeerDescuentosAgotados(string path)
        {
            var listDescuentosAgotados = new List<DescuentosAgotados>();

            using (var reader = new StreamReader(path))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    var descuentoAgotado = new DescuentosAgotados(
                        Functions.ConvertToInt(values[0]),
                        values[1],
                        values[2]
                    );

                    listDescuentosAgotados.Add(descuentoAgotado);
                }
            }

            return listDescuentosAgotados;
        }

        public static List<QueryMaestra> LeerQueryMaestra(string path)
        {
            int pos = 0;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfworkbook.GetSheet("Sheet");

            try
            {
                DataFormatter formatter = new DataFormatter();
                var listDescuentosNuevos = new List<QueryMaestra>();

                for (int row = 0; row <= sheet.LastRowNum; row++)
                {
                    if (row == 0) continue;
                    pos++;

                    if (sheet.GetRow(row) != null)
                    {
                        var descuentoNuevo = new QueryMaestra(
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(0))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(1))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(2))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(3))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(4))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(5))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(6))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(7))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(8))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(9))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(10))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(11))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(12))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(13))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(14))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(15))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(16))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(17))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(18))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(19))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(20))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(21))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(22))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(23))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(24))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(25))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(26))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(27))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(28))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(29))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(30))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(31))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(32))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(33))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(34))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(35))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(36))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(37))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(38))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(39))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(40))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(41))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(42))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(43))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(44))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(45))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(46))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(47))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(48))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(49)))
                        );

                        listDescuentosNuevos.Add(descuentoNuevo);
                    }
                }

                
                return listDescuentosNuevos;
            }
            catch(Exception ex)
            {
                Console.WriteLine("Ha ocurrido un error en el registro {0}. El error es el siguiente: {1}", pos, ex.Message);
                return null;
            }
        }

        public static List<QueryMaestra> LeerQueryMaestraTSV(string path)
        {
            //var engine = new FileHelperAsyncEngine<QueryMaestra>();
            var engine = new DelimitedFileEngine<QueryMaestra>();

            engine.ErrorManager.ErrorMode = ErrorMode.IgnoreAndContinue;

            var resultados = new List<QueryMaestra>();

            try
            {
                var results = engine.ReadFile(path);

                foreach (var err in engine.ErrorManager.Errors)
                {
                    Console.WriteLine();
                    Console.WriteLine("Error on Line number: {0}", err.LineNumber);
                    Console.WriteLine("Record causing the problem: {0}", err.RecordString);
                    Console.WriteLine("Complete exception information: {0}", err.ExceptionInfo.ToString());
                }

                foreach (var res in results)
                {
                    if (!string.IsNullOrEmpty(res.FechaFin))
                    {
                        //res.FechaFin = Functions.ConvertToDateFromRegexDate(res.FechaFin);
                    }
                    resultados.Add(res);
                }
            }
            catch (Exception ex)
            {
                
                Console.WriteLine(ex.Message);
            }

            //Elimina los datos de la cabecera
            resultados.RemoveAt(0);

            return resultados;
        }

        public static MaestraTotal LeerMaestraCompleta(string path)
        {
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new XSSFWorkbook(file);
            }


            DataFormatter formatter = new DataFormatter();
            var maestraCompleta = new MaestraTotal();

            try
            {
                maestraCompleta.MaestraDescuentos = ObtenerMaestraDescuentos(formatter, "MAESTRA_DESCUENTOS_V1");
                maestraCompleta.MaestraFuentes = ObtenerMaestraFuentes(formatter, "MAESTRA_Fuentes");
                maestraCompleta.MaestraEstablecimientos = ObtenerMaestraEstablecimientos(formatter, "MAESTRA_Establecimientos");

                return maestraCompleta;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ha ocurrido un error en el registro. El error es el siguiente: {0}", ex.Message);
                return null;
            }
        }

        public static List<MaestraDescuentos> ObtenerMaestraDescuentos(DataFormatter formatter, string sheetName)
        {
            var maestraDescuentos = new List<MaestraDescuentos>();

            ISheet sheetDescuentos = hssfworkbook.GetSheet(sheetName);
            for (int row = 0; row <= sheetDescuentos.LastRowNum; row++)
            {
                if (row == 0) continue;

                if (sheetDescuentos.GetRow(row) != null)
                {
                    var descuento = new MaestraDescuentos(
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(0))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(1))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(2))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(3))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(4))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(5))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(6))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(7))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(8))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(9))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(10))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(11))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(12))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(13))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(14))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(15))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(16))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(17))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(18))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(19))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(20))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(21))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(22))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(23))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(24))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(25))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(26))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(27))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(28))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(29))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(30))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(31))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(32))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(33))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(34))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(35))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(36))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(37))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(38))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(39))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(40))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(41))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetDescuentos.GetRow(row).GetCell(42)))
                    );

                    maestraDescuentos.Add(descuento);
                }
            }

            return maestraDescuentos;
        }

        public static List<MaestraFuentes> ObtenerMaestraFuentes(DataFormatter formatter, string sheetName)
        {
            var maestraFuentes = new List<MaestraFuentes>();

            ISheet sheetFuentes = hssfworkbook.GetSheet(sheetName);
            for (int row = 0; row <= sheetFuentes.LastRowNum; row++)
            {
                if (row == 0) continue;

                if (sheetFuentes.GetRow(row) != null)
                {
                    var fuente = new MaestraFuentes(
                        Functions.ConvertToString(formatter.FormatCellValue(sheetFuentes.GetRow(row).GetCell(0))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetFuentes.GetRow(row).GetCell(1))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetFuentes.GetRow(row).GetCell(2))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetFuentes.GetRow(row).GetCell(3)))
                    );

                    maestraFuentes.Add(fuente);
                }
            }

            return maestraFuentes;
        }

        public static List<MaestraEstablecimientos> ObtenerMaestraEstablecimientos(DataFormatter formatter, string sheetName)
        {
            var maestraEstablecimientos = new List<MaestraEstablecimientos>();

            ISheet sheetEstablecimientos = hssfworkbook.GetSheet(sheetName);
            for (int row = 0; row <= sheetEstablecimientos.LastRowNum; row++)
            {
                if (row == 0) continue;

                if (sheetEstablecimientos.GetRow(row) != null)
                {
                    var establecimiento = new MaestraEstablecimientos(
                        Functions.ConvertToString(formatter.FormatCellValue(sheetEstablecimientos.GetRow(row).GetCell(0))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetEstablecimientos.GetRow(row).GetCell(1))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetEstablecimientos.GetRow(row).GetCell(2))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetEstablecimientos.GetRow(row).GetCell(3))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetEstablecimientos.GetRow(row).GetCell(3))),
                        Functions.ConvertToString(formatter.FormatCellValue(sheetEstablecimientos.GetRow(row).GetCell(3)))
                    );

                    maestraEstablecimientos.Add(establecimiento);
                }
            }

            return maestraEstablecimientos;
        }

        public static void ExportarMaestraModificada(MaestraTotal maestraTotal, string pathSalida, string fileName)
        {
            hssfworkbook = new XSSFWorkbook();

            try
            {
                #region MAESTRA_DESCUENTOS_V1
                ISheet sheetV1 = hssfworkbook.CreateSheet("MAESTRA_DESCUENTOS_V1");
                IRow rowHeaderV1 = sheetV1.CreateRow(0);
                rowHeaderV1.CreateCell(0).SetCellValue("idGrupoBeneficio");
                rowHeaderV1.CreateCell(1).SetCellValue("idBeneficio");
                rowHeaderV1.CreateCell(2).SetCellValue("fuenteDescuentoGrupo");
                rowHeaderV1.CreateCell(3).SetCellValue("idFuenteDescuento");
                rowHeaderV1.CreateCell(4).SetCellValue("fuenteDescuento");
                rowHeaderV1.CreateCell(5).SetCellValue("flagEstado");
                rowHeaderV1.CreateCell(6).SetCellValue("imgDescuento");
                rowHeaderV1.CreateCell(7).SetCellValue("imgLogoEmpresa");
                rowHeaderV1.CreateCell(8).SetCellValue("imgLogoFuente");
                rowHeaderV1.CreateCell(9).SetCellValue("tituloBeneficio");
                rowHeaderV1.CreateCell(10).SetCellValue("idEstablecimiento");
                rowHeaderV1.CreateCell(11).SetCellValue("nombreEstablecimiento");
                rowHeaderV1.CreateCell(12).SetCellValue("porcentaje");
                rowHeaderV1.CreateCell(13).SetCellValue("precioFinal");
                rowHeaderV1.CreateCell(14).SetCellValue("moneda");
                rowHeaderV1.CreateCell(15).SetCellValue("descripcionResumen");
                rowHeaderV1.CreateCell(16).SetCellValue("fechaInicio");
                rowHeaderV1.CreateCell(17).SetCellValue("fechaFin");
                rowHeaderV1.CreateCell(18).SetCellValue("detalleBeneficio");
                rowHeaderV1.CreateCell(19).SetCellValue("restricciones");
                rowHeaderV1.CreateCell(20).SetCellValue("detalleAdicional");
                rowHeaderV1.CreateCell(21).SetCellValue("diasAtencionEstab");
                rowHeaderV1.CreateCell(22).SetCellValue("horarioEstablec");
                rowHeaderV1.CreateCell(23).SetCellValue("ruc");
                rowHeaderV1.CreateCell(24).SetCellValue("razonSocial");
                rowHeaderV1.CreateCell(25).SetCellValue("direccion");
                rowHeaderV1.CreateCell(26).SetCellValue("latitud");
                rowHeaderV1.CreateCell(27).SetCellValue("longitud");
                rowHeaderV1.CreateCell(28).SetCellValue("flagOnline");
                rowHeaderV1.CreateCell(29).SetCellValue("rubroBcpApp");
                rowHeaderV1.CreateCell(30).SetCellValue("rubroParati");
                rowHeaderV1.CreateCell(31).SetCellValue("rubroParatiApp");
                rowHeaderV1.CreateCell(32).SetCellValue("departamentoEstablec");
                rowHeaderV1.CreateCell(33).SetCellValue("provinciaEstablec");
                rowHeaderV1.CreateCell(34).SetCellValue("distritoEstablec");
                rowHeaderV1.CreateCell(35).SetCellValue("segmentoCliente");
                rowHeaderV1.CreateCell(36).SetCellValue("url");
                rowHeaderV1.CreateCell(37).SetCellValue("sexo");
                rowHeaderV1.CreateCell(38).SetCellValue("mensajeMotivador");
                rowHeaderV1.CreateCell(39).SetCellValue("tipoSuscripcion");
                rowHeaderV1.CreateCell(40).SetCellValue("tiporigen");
                rowHeaderV1.CreateCell(41).SetCellValue("codpromocion");
                rowHeaderV1.CreateCell(42).SetCellValue("fechaProceso");


                for (int i = 0; i < maestraTotal.MaestraDescuentos.Count; i++)
                {
                    IRow rowDataV1 = sheetV1.CreateRow(i + 1);
                    bool parseIdGrupoBeneficio = float.TryParse(maestraTotal.MaestraDescuentos[i].IdGrupoBeneficio, out float IdBrupoBeneficio);
                    if (parseIdGrupoBeneficio) rowDataV1.CreateCell(0).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].IdGrupoBeneficio).Value);
                    else rowDataV1.CreateCell(0).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].IdGrupoBeneficio));


                    bool parseIdBeneficio = float.TryParse(maestraTotal.MaestraDescuentos[i].IdBeneficio, out float IdBeneficio);
                    if (parseIdBeneficio) rowDataV1.CreateCell(1).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].IdBeneficio).Value);
                    else rowDataV1.CreateCell(1).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].IdBeneficio));


                    bool parseFuenteDescuentoGrupo = float.TryParse(maestraTotal.MaestraDescuentos[i].FuenteDescuentoGrupo, out float FuenteDescuentoGrupo);
                    if (parseFuenteDescuentoGrupo) rowDataV1.CreateCell(2).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].FuenteDescuentoGrupo).Value);
                    else rowDataV1.CreateCell(2).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].FuenteDescuentoGrupo));


                    bool parseIdFuenteDescuento = float.TryParse(maestraTotal.MaestraDescuentos[i].IdFuenteDescuento, out float IdFuenteDescuento);
                    if (parseIdFuenteDescuento) rowDataV1.CreateCell(3).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].IdFuenteDescuento).Value);
                    else rowDataV1.CreateCell(3).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].IdFuenteDescuento));


                    bool parseFuenteDescuento = float.TryParse(maestraTotal.MaestraDescuentos[i].FuenteDescuento, out float FuenteDescuento);
                    if (parseFuenteDescuento) rowDataV1.CreateCell(4).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].FuenteDescuento).Value);
                    else rowDataV1.CreateCell(4).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].FuenteDescuento));


                    bool parseFlagEstado = float.TryParse(maestraTotal.MaestraDescuentos[i].FlagEstado, out float FlagEstado);
                    if (parseFlagEstado) rowDataV1.CreateCell(5).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].FlagEstado).Value);
                    else rowDataV1.CreateCell(5).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].FlagEstado));


                    bool parseImgDescuento = float.TryParse(maestraTotal.MaestraDescuentos[i].ImgDescuento, out float ImgDescuento);
                    if (parseImgDescuento) rowDataV1.CreateCell(6).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].ImgDescuento).Value);
                    else rowDataV1.CreateCell(6).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].ImgDescuento));


                    bool parseImgLogoEmpresa = float.TryParse(maestraTotal.MaestraDescuentos[i].ImgLogoEmpresa, out float ImgLogoEmpresa);
                    if (parseImgLogoEmpresa) rowDataV1.CreateCell(7).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].ImgLogoEmpresa).Value);
                    else rowDataV1.CreateCell(7).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].ImgLogoEmpresa));


                    bool parseImgLogoFuente = float.TryParse(maestraTotal.MaestraDescuentos[i].ImgLogoFuente, out float ImgLogoFuente);
                    if (parseImgLogoFuente) rowDataV1.CreateCell(8).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].ImgLogoFuente).Value);
                    else rowDataV1.CreateCell(8).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].ImgLogoFuente));


                    bool parseTituloBeneficio = float.TryParse(maestraTotal.MaestraDescuentos[i].TituloBeneficio, out float TituloBeneficio);
                    if (parseTituloBeneficio) rowDataV1.CreateCell(9).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].TituloBeneficio).Value);
                    else rowDataV1.CreateCell(9).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].TituloBeneficio));


                    bool parseIdEstablecimiento = float.TryParse(maestraTotal.MaestraDescuentos[i].IdEstablecimiento, out float IdEstablecimiento);
                    if (parseIdEstablecimiento) rowDataV1.CreateCell(10).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].IdEstablecimiento).Value);
                    else rowDataV1.CreateCell(10).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].IdEstablecimiento));


                    bool parseNombreEstablecimiento = float.TryParse(maestraTotal.MaestraDescuentos[i].NombreEstablecimiento, out float NombreEstablecimiento);
                    if (parseNombreEstablecimiento) rowDataV1.CreateCell(11).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].NombreEstablecimiento).Value);
                    else rowDataV1.CreateCell(11).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].NombreEstablecimiento));


                    bool parsePorcentaje = float.TryParse(maestraTotal.MaestraDescuentos[i].Porcentaje, out float Porcentaje);
                    if (parsePorcentaje) rowDataV1.CreateCell(12).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].Porcentaje).Value);
                    else rowDataV1.CreateCell(12).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].Porcentaje));


                    bool parsePrecioFinal = float.TryParse(maestraTotal.MaestraDescuentos[i].PrecioFinal, out float PrecioFinal);
                    if (parsePrecioFinal) rowDataV1.CreateCell(13).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].PrecioFinal).Value);
                    else rowDataV1.CreateCell(13).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].PrecioFinal));


                    bool parseMoneda = float.TryParse(maestraTotal.MaestraDescuentos[i].Moneda, out float Moneda);
                    if (parseMoneda) rowDataV1.CreateCell(14).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].Moneda).Value);
                    else rowDataV1.CreateCell(14).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].Moneda));


                    bool parseDescripcionResumen = float.TryParse(maestraTotal.MaestraDescuentos[i].DescripcionResumen, out float DescripcionResumen);
                    if (parseDescripcionResumen) rowDataV1.CreateCell(15).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].DescripcionResumen).Value);
                    else rowDataV1.CreateCell(15).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].DescripcionResumen));


                    bool parseFechaInicio = float.TryParse(maestraTotal.MaestraDescuentos[i].FechaInicio, out float FechaInicio);
                    if (parseFechaInicio) rowDataV1.CreateCell(16).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].FechaInicio).Value);
                    else rowDataV1.CreateCell(16).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].FechaInicio));


                    bool parseFechaFin = float.TryParse(maestraTotal.MaestraDescuentos[i].FechaFin, out float FechaFin);
                    if (parseFechaFin) rowDataV1.CreateCell(17).SetCellValue(maestraTotal.MaestraDescuentos[i].FechaFin);
                    else rowDataV1.CreateCell(17).SetCellValue(maestraTotal.MaestraDescuentos[i].FechaFin);


                    bool parseDetalleBeneficio = float.TryParse(maestraTotal.MaestraDescuentos[i].DetalleBeneficio, out float DetalleBeneficio);
                    if(parseDetalleBeneficio) rowDataV1.CreateCell(18).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].DetalleBeneficio).Value);
                    else rowDataV1.CreateCell(18).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].DetalleBeneficio));


                    bool parseRestricciones = float.TryParse(maestraTotal.MaestraDescuentos[i].Restricciones, out float Restricciones);
                    if (parseRestricciones) rowDataV1.CreateCell(19).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].Restricciones).Value);
                    else rowDataV1.CreateCell(19).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].Restricciones));


                    bool parseDetalleAdicional = float.TryParse(maestraTotal.MaestraDescuentos[i].DetalleAdicional, out float DetalleAdicional);
                    if (parseDetalleAdicional) rowDataV1.CreateCell(20).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].DetalleAdicional).Value);
                    else rowDataV1.CreateCell(20).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].DetalleAdicional));


                    bool parseDiasAtencionEstab = float.TryParse(maestraTotal.MaestraDescuentos[i].DiasAtencionEstab, out float DiasAtencionEstab);
                    if (parseDiasAtencionEstab) rowDataV1.CreateCell(21).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].DiasAtencionEstab).Value);
                    else rowDataV1.CreateCell(21).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].DiasAtencionEstab));


                    bool parseHorarioEstablec = float.TryParse(maestraTotal.MaestraDescuentos[i].HorarioEstablec, out float HorarioEstablec);
                    if (parseHorarioEstablec) rowDataV1.CreateCell(22).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].HorarioEstablec).Value);
                    else rowDataV1.CreateCell(22).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].HorarioEstablec));


                    bool parseRuc = float.TryParse(maestraTotal.MaestraDescuentos[i].Ruc, out float Ruc);
                    if (parseRuc) rowDataV1.CreateCell(23).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].Ruc).Value);
                    else rowDataV1.CreateCell(23).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].Ruc));


                    bool parseRazonSocial = float.TryParse(maestraTotal.MaestraDescuentos[i].RazonSocial, out float RazonSocial);
                    if (parseRazonSocial) rowDataV1.CreateCell(24).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].RazonSocial).Value);
                    else rowDataV1.CreateCell(24).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].RazonSocial));


                    bool parseDireccion = float.TryParse(maestraTotal.MaestraDescuentos[i].Direccion, out float Direccion);
                    if (parseDireccion) rowDataV1.CreateCell(25).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].Direccion).Value);
                    else rowDataV1.CreateCell(25).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].Direccion));


                    bool parseLatitud = float.TryParse(maestraTotal.MaestraDescuentos[i].Latitud, out float Latitud);
                    if (parseLatitud) rowDataV1.CreateCell(26).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].Latitud).Value);
                    else rowDataV1.CreateCell(26).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].Latitud));


                    bool parseLongitud = float.TryParse(maestraTotal.MaestraDescuentos[i].Longitud, out float Longitud);
                    if (parseLongitud) rowDataV1.CreateCell(27).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].Longitud).Value);
                    else rowDataV1.CreateCell(27).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].Longitud));


                    bool parseFlagOnline = float.TryParse(maestraTotal.MaestraDescuentos[i].FlagOnline, out float FlagOnline);
                    if (parseFlagOnline) rowDataV1.CreateCell(28).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].FlagOnline).Value);
                    else rowDataV1.CreateCell(28).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].FlagOnline));


                    bool parseRubroBcpApp = float.TryParse(maestraTotal.MaestraDescuentos[i].RubroBcpApp, out float RubroBcpApp);
                    if (parseRubroBcpApp) rowDataV1.CreateCell(29).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].RubroBcpApp).Value);
                    else rowDataV1.CreateCell(29).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].RubroBcpApp));


                    bool parseRubroParaTi = float.TryParse(maestraTotal.MaestraDescuentos[i].RubroParaTi, out float RubroParaTi);
                    if (parseRubroParaTi) rowDataV1.CreateCell(30).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].RubroParaTi).Value);
                    else rowDataV1.CreateCell(30).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].RubroParaTi));


                    bool parseRubroParaTiApp = float.TryParse(maestraTotal.MaestraDescuentos[i].RubroParaTiApp, out float RubroParaTiApp);
                    if (parseRubroParaTiApp) rowDataV1.CreateCell(31).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].RubroParaTiApp).Value);
                    else rowDataV1.CreateCell(31).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].RubroParaTiApp));


                    bool parseDepartamentoEstablec = float.TryParse(maestraTotal.MaestraDescuentos[i].DepartamentoEstablec, out float DepartamentoEstablec);
                    if (parseDepartamentoEstablec) rowDataV1.CreateCell(32).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].DepartamentoEstablec).Value);
                    else rowDataV1.CreateCell(32).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].DepartamentoEstablec));


                    bool parseProvinciaEstablec = float.TryParse(maestraTotal.MaestraDescuentos[i].ProvinciaEstablec, out float ProvinciaEstablec);
                    if (parseProvinciaEstablec) rowDataV1.CreateCell(33).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].ProvinciaEstablec).Value);
                    else rowDataV1.CreateCell(33).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].ProvinciaEstablec));


                    bool parseDistritoEstablec = float.TryParse(maestraTotal.MaestraDescuentos[i].DistritoEstablec, out float DistritoEstablec);
                    if (parseDistritoEstablec) rowDataV1.CreateCell(34).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].DistritoEstablec).Value);
                    else rowDataV1.CreateCell(34).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].DistritoEstablec));


                    bool parseSegmentoCliente = float.TryParse(maestraTotal.MaestraDescuentos[i].SegmentoCliente, out float SegmentoCliente);
                    if (parseSegmentoCliente) rowDataV1.CreateCell(35).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].SegmentoCliente).Value);
                    else rowDataV1.CreateCell(35).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].SegmentoCliente));


                    bool parseUrl = float.TryParse(maestraTotal.MaestraDescuentos[i].Url, out float Url);
                    if (parseUrl) rowDataV1.CreateCell(36).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].Url).Value);
                    else rowDataV1.CreateCell(36).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].Url));


                    bool parseSexo = float.TryParse(maestraTotal.MaestraDescuentos[i].Sexo, out float Sexo);
                    if (parseSexo) rowDataV1.CreateCell(37).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].Sexo).Value);
                    else rowDataV1.CreateCell(37).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].Sexo));


                    bool parseMensajeMotivador = float.TryParse(maestraTotal.MaestraDescuentos[i].MensajeMotivador, out float MensajeMotivador);
                    if (parseMensajeMotivador) rowDataV1.CreateCell(38).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].MensajeMotivador).Value);
                    else rowDataV1.CreateCell(38).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].MensajeMotivador));


                    bool parseTipoSubscripcion = float.TryParse(maestraTotal.MaestraDescuentos[i].TipoSubscripcion, out float TipoSubscripcion);
                    if (parseTipoSubscripcion) rowDataV1.CreateCell(39).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].TipoSubscripcion).Value);
                    else rowDataV1.CreateCell(39).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].TipoSubscripcion));


                    bool parseTipoOrigen = float.TryParse(maestraTotal.MaestraDescuentos[i].TipoOrigen, out float TipoOrigen);
                    if (parseTipoOrigen) rowDataV1.CreateCell(40).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].TipoOrigen).Value);
                    else rowDataV1.CreateCell(40).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].TipoOrigen));


                    bool parseCodPromocion = float.TryParse(maestraTotal.MaestraDescuentos[i].CodPromocion, out float CodPromocion);
                    if (parseCodPromocion) rowDataV1.CreateCell(41).SetCellValue(Functions.ConvertToFloat(maestraTotal.MaestraDescuentos[i].CodPromocion).Value);
                    else rowDataV1.CreateCell(41).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraDescuentos[i].CodPromocion));


                    bool parseFechaProceso = float.TryParse(maestraTotal.MaestraDescuentos[i].FechaProceso, out float FechaProceso);
                    if (parseFechaProceso) rowDataV1.CreateCell(42).SetCellValue(maestraTotal.MaestraDescuentos[i].FechaProceso);
                    else rowDataV1.CreateCell(42).SetCellValue(maestraTotal.MaestraDescuentos[i].FechaProceso);

                }
                #endregion MAESTRA_DESCUENTOS_V1


                #region MAESTRA_FUENTES
                ISheet sheetV2 = hssfworkbook.CreateSheet("MAESTRA_Fuentes");
                IRow rowHeaderV2 = sheetV2.CreateRow(0);
                rowHeaderV2.CreateCell(0).SetCellValue("fuenteDescuento");
                rowHeaderV2.CreateCell(1).SetCellValue("imgLogoFuente");
                rowHeaderV2.CreateCell(2).SetCellValue("idFuenteDescuento");
                rowHeaderV2.CreateCell(3).SetCellValue("tipoSuscripcion");

                for (int i = 0; i < maestraTotal.MaestraFuentes.Count; i++)
                {
                    IRow rowDataV2 = sheetV2.CreateRow(i + 1);
                    rowDataV2.CreateCell(0).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraFuentes[i].FuenteDescuento));
                    rowDataV2.CreateCell(1).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraFuentes[i].ImgLogoFuente));
                    rowDataV2.CreateCell(2).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraFuentes[i].IdFuenteDescuento));
                    rowDataV2.CreateCell(3).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraFuentes[i].TipoSuscripcion));
                }
                #endregion MAESTRA_FUENTES


                #region MAESTRA_ESTABLECIMIENTOS
                ISheet sheetV3 = hssfworkbook.CreateSheet("MAESTRA_Establecimientos");
                IRow rowHeaderV3 = sheetV3.CreateRow(0);
                rowHeaderV3.CreateCell(0).SetCellValue("IdEstablecimiento");
                rowHeaderV3.CreateCell(1).SetCellValue("nombreEstablecimiento");
                rowHeaderV3.CreateCell(2).SetCellValue("formato");
                rowHeaderV3.CreateCell(3).SetCellValue("imgLogoEmpresa");
                rowHeaderV3.CreateCell(2).SetCellValue("observacion");
                rowHeaderV3.CreateCell(3).SetCellValue("imgLogoEmpresa_Old");

                for (int i = 0; i < maestraTotal.MaestraEstablecimientos.Count; i++)
                {
                    IRow rowDataV3 = sheetV3.CreateRow(i + 1);
                    rowDataV3.CreateCell(0).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraEstablecimientos[i].IdEstablecimiento));
                    rowDataV3.CreateCell(1).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraEstablecimientos[i].NombreEstablecimiento));
                    rowDataV3.CreateCell(2).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraEstablecimientos[i].Formato));
                    rowDataV3.CreateCell(3).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraEstablecimientos[i].ImgLogoEmpresa));
                    rowDataV3.CreateCell(2).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraEstablecimientos[i].Observacion));
                    rowDataV3.CreateCell(3).SetCellValue(Functions.ConvertToString(maestraTotal.MaestraEstablecimientos[i].ImgLogoEmpresaOld));
                }
                #endregion MAESTRA_ESTABLECIMIENTOS

                string filePath = $"{pathSalida}/{fileName}";
                using (FileStream file = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write))
                {
                    hssfworkbook.Write(file);
                    file.Close();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("El error es el siguiente: {0}", ex.Message);
            }

            hssfworkbook = null;
        }

        public static List<DeEmpresa> LeerDeEmpresaExcel(string path)
        {
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfworkbook.GetSheet("Hoja1");

            try
            {
                DataFormatter formatter = new DataFormatter();
                var listDeEmpresa = new List<DeEmpresa>();

                for (int row = 0; row <= sheet.LastRowNum; row++)
                {
                    if (row == 0) continue;

                    if (sheet.GetRow(row) != null)
                    {
                        var deEmpresa = new DeEmpresa(
                            (row + 1).ToString(),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(0))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(1))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(2))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(3))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(4))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(5))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(6))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(7))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(8))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(9))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(10))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(11))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(12))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(13))),
                            Functions.ConvertToString(formatter.FormatCellValue(sheet.GetRow(row).GetCell(14)))
                        );

                        listDeEmpresa.Add(deEmpresa);
                    }
                }

                return listDeEmpresa;
            }
            catch(Exception ex)
            {
                Console.WriteLine("Ha ocurrido un error. El error es el siguiente: {0}", ex.Message);
                return null;
            }
        }

        public static void ExportarRechazados(List<MaestraDescuentos> maestraDescuento, string pathSalida, string fileName)
        {
            hssfworkbook = new XSSFWorkbook();
            try
            {
                ISheet sheetV1 = hssfworkbook.CreateSheet("Rechazados");
                IRow rowHeaderV1 = sheetV1.CreateRow(0);
                rowHeaderV1.CreateCell(0).SetCellValue("idGrupoBeneficio");
                rowHeaderV1.CreateCell(1).SetCellValue("idBeneficio");
                rowHeaderV1.CreateCell(2).SetCellValue("fuenteDescuentoGrupo");
                rowHeaderV1.CreateCell(3).SetCellValue("idFuenteDescuento");
                rowHeaderV1.CreateCell(4).SetCellValue("fuenteDescuento");
                rowHeaderV1.CreateCell(5).SetCellValue("flagEstado");
                rowHeaderV1.CreateCell(6).SetCellValue("imgDescuento");
                rowHeaderV1.CreateCell(7).SetCellValue("imgLogoEmpresa");
                rowHeaderV1.CreateCell(8).SetCellValue("imgLogoFuente");
                rowHeaderV1.CreateCell(9).SetCellValue("tituloBeneficio");
                rowHeaderV1.CreateCell(10).SetCellValue("idEstablecimiento");
                rowHeaderV1.CreateCell(11).SetCellValue("nombreEstablecimiento");
                rowHeaderV1.CreateCell(12).SetCellValue("porcentaje");
                rowHeaderV1.CreateCell(13).SetCellValue("precioFinal");
                rowHeaderV1.CreateCell(14).SetCellValue("moneda");
                rowHeaderV1.CreateCell(15).SetCellValue("descripcionResumen");
                rowHeaderV1.CreateCell(16).SetCellValue("fechaInicio");
                rowHeaderV1.CreateCell(17).SetCellValue("fechaFin");
                rowHeaderV1.CreateCell(18).SetCellValue("detalleBeneficio");
                rowHeaderV1.CreateCell(19).SetCellValue("restricciones");
                rowHeaderV1.CreateCell(20).SetCellValue("detalleAdicional");
                rowHeaderV1.CreateCell(21).SetCellValue("diasAtencionEstab");
                rowHeaderV1.CreateCell(22).SetCellValue("horarioEstablec");
                rowHeaderV1.CreateCell(23).SetCellValue("ruc");
                rowHeaderV1.CreateCell(24).SetCellValue("razonSocial");
                rowHeaderV1.CreateCell(25).SetCellValue("direccion");
                rowHeaderV1.CreateCell(26).SetCellValue("latitud");
                rowHeaderV1.CreateCell(27).SetCellValue("longitud");
                rowHeaderV1.CreateCell(28).SetCellValue("flagOnline");
                rowHeaderV1.CreateCell(29).SetCellValue("rubroBcpApp");
                rowHeaderV1.CreateCell(30).SetCellValue("rubroParati");
                rowHeaderV1.CreateCell(31).SetCellValue("rubroParatiApp");
                rowHeaderV1.CreateCell(32).SetCellValue("departamentoEstablec");
                rowHeaderV1.CreateCell(33).SetCellValue("provinciaEstablec");
                rowHeaderV1.CreateCell(34).SetCellValue("distritoEstablec");
                rowHeaderV1.CreateCell(35).SetCellValue("segmentoCliente");
                rowHeaderV1.CreateCell(36).SetCellValue("url");
                rowHeaderV1.CreateCell(37).SetCellValue("sexo");
                rowHeaderV1.CreateCell(38).SetCellValue("mensajeMotivador");
                rowHeaderV1.CreateCell(39).SetCellValue("tipoSuscripcion");
                rowHeaderV1.CreateCell(40).SetCellValue("tiporigen");
                rowHeaderV1.CreateCell(41).SetCellValue("codpromocion");
                rowHeaderV1.CreateCell(42).SetCellValue("fechaProceso");


                for (int i = 0; i < maestraDescuento.Count; i++)
                {
                    IRow rowDataV1 = sheetV1.CreateRow(i + 1);
                    bool parseIdGrupoBeneficio = float.TryParse(maestraDescuento[i].IdGrupoBeneficio, out float IdBrupoBeneficio);
                    if (parseIdGrupoBeneficio) rowDataV1.CreateCell(0).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].IdGrupoBeneficio).Value);
                    else rowDataV1.CreateCell(0).SetCellValue(Functions.ConvertToString(maestraDescuento[i].IdGrupoBeneficio));


                    bool parseIdBeneficio = float.TryParse(maestraDescuento[i].IdBeneficio, out float IdBeneficio);
                    if (parseIdBeneficio) rowDataV1.CreateCell(1).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].IdBeneficio).Value);
                    else rowDataV1.CreateCell(1).SetCellValue(Functions.ConvertToString(maestraDescuento[i].IdBeneficio));


                    bool parseFuenteDescuentoGrupo = float.TryParse(maestraDescuento[i].FuenteDescuentoGrupo, out float FuenteDescuentoGrupo);
                    if (parseFuenteDescuentoGrupo) rowDataV1.CreateCell(2).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].FuenteDescuentoGrupo).Value);
                    else rowDataV1.CreateCell(2).SetCellValue(Functions.ConvertToString(maestraDescuento[i].FuenteDescuentoGrupo));


                    bool parseIdFuenteDescuento = float.TryParse(maestraDescuento[i].IdFuenteDescuento, out float IdFuenteDescuento);
                    if (parseIdFuenteDescuento) rowDataV1.CreateCell(3).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].IdFuenteDescuento).Value);
                    else rowDataV1.CreateCell(3).SetCellValue(Functions.ConvertToString(maestraDescuento[i].IdFuenteDescuento));


                    bool parseFuenteDescuento = float.TryParse(maestraDescuento[i].FuenteDescuento, out float FuenteDescuento);
                    if (parseFuenteDescuento) rowDataV1.CreateCell(4).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].FuenteDescuento).Value);
                    else rowDataV1.CreateCell(4).SetCellValue(Functions.ConvertToString(maestraDescuento[i].FuenteDescuento));


                    bool parseFlagEstado = float.TryParse(maestraDescuento[i].FlagEstado, out float FlagEstado);
                    if (parseFlagEstado) rowDataV1.CreateCell(5).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].FlagEstado).Value);
                    else rowDataV1.CreateCell(5).SetCellValue(Functions.ConvertToString(maestraDescuento[i].FlagEstado));


                    bool parseImgDescuento = float.TryParse(maestraDescuento[i].ImgDescuento, out float ImgDescuento);
                    if (parseImgDescuento) rowDataV1.CreateCell(6).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].ImgDescuento).Value);
                    else rowDataV1.CreateCell(6).SetCellValue(Functions.ConvertToString(maestraDescuento[i].ImgDescuento));


                    bool parseImgLogoEmpresa = float.TryParse(maestraDescuento[i].ImgLogoEmpresa, out float ImgLogoEmpresa);
                    if (parseImgLogoEmpresa) rowDataV1.CreateCell(7).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].ImgLogoEmpresa).Value);
                    else rowDataV1.CreateCell(7).SetCellValue(Functions.ConvertToString(maestraDescuento[i].ImgLogoEmpresa));


                    bool parseImgLogoFuente = float.TryParse(maestraDescuento[i].ImgLogoFuente, out float ImgLogoFuente);
                    if (parseImgLogoFuente) rowDataV1.CreateCell(8).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].ImgLogoFuente).Value);
                    else rowDataV1.CreateCell(8).SetCellValue(Functions.ConvertToString(maestraDescuento[i].ImgLogoFuente));


                    bool parseTituloBeneficio = float.TryParse(maestraDescuento[i].TituloBeneficio, out float TituloBeneficio);
                    if (parseTituloBeneficio) rowDataV1.CreateCell(9).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].TituloBeneficio).Value);
                    else rowDataV1.CreateCell(9).SetCellValue(Functions.ConvertToString(maestraDescuento[i].TituloBeneficio));


                    bool parseIdEstablecimiento = float.TryParse(maestraDescuento[i].IdEstablecimiento, out float IdEstablecimiento);
                    if (parseIdEstablecimiento) rowDataV1.CreateCell(10).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].IdEstablecimiento).Value);
                    else rowDataV1.CreateCell(10).SetCellValue(Functions.ConvertToString(maestraDescuento[i].IdEstablecimiento));


                    bool parseNombreEstablecimiento = float.TryParse(maestraDescuento[i].NombreEstablecimiento, out float NombreEstablecimiento);
                    if (parseNombreEstablecimiento) rowDataV1.CreateCell(11).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].NombreEstablecimiento).Value);
                    else rowDataV1.CreateCell(11).SetCellValue(Functions.ConvertToString(maestraDescuento[i].NombreEstablecimiento));


                    bool parsePorcentaje = float.TryParse(maestraDescuento[i].Porcentaje, out float Porcentaje);
                    if (parsePorcentaje) rowDataV1.CreateCell(12).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].Porcentaje).Value);
                    else rowDataV1.CreateCell(12).SetCellValue(Functions.ConvertToString(maestraDescuento[i].Porcentaje));


                    bool parsePrecioFinal = float.TryParse(maestraDescuento[i].PrecioFinal, out float PrecioFinal);
                    if (parsePrecioFinal) rowDataV1.CreateCell(13).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].PrecioFinal).Value);
                    else rowDataV1.CreateCell(13).SetCellValue(Functions.ConvertToString(maestraDescuento[i].PrecioFinal));


                    bool parseMoneda = float.TryParse(maestraDescuento[i].Moneda, out float Moneda);
                    if (parseMoneda) rowDataV1.CreateCell(14).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].Moneda).Value);
                    else rowDataV1.CreateCell(14).SetCellValue(Functions.ConvertToString(maestraDescuento[i].Moneda));


                    bool parseDescripcionResumen = float.TryParse(maestraDescuento[i].DescripcionResumen, out float DescripcionResumen);
                    if (parseDescripcionResumen) rowDataV1.CreateCell(15).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].DescripcionResumen).Value);
                    else rowDataV1.CreateCell(15).SetCellValue(Functions.ConvertToString(maestraDescuento[i].DescripcionResumen));


                    bool parseFechaInicio = float.TryParse(maestraDescuento[i].FechaInicio, out float FechaInicio);
                    if (parseFechaInicio) rowDataV1.CreateCell(16).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].FechaInicio).Value);
                    else rowDataV1.CreateCell(16).SetCellValue(Functions.ConvertToString(maestraDescuento[i].FechaInicio));


                    bool parseFechaFin = float.TryParse(maestraDescuento[i].FechaFin, out float FechaFin);
                    if (parseFechaFin) rowDataV1.CreateCell(17).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].FechaFin).Value);
                    else rowDataV1.CreateCell(17).SetCellValue(Functions.ConvertToString(maestraDescuento[i].FechaFin));


                    bool parseDetalleBeneficio = float.TryParse(maestraDescuento[i].DetalleBeneficio, out float DetalleBeneficio);
                    if(parseDetalleBeneficio) rowDataV1.CreateCell(18).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].DetalleBeneficio).Value);
                    else rowDataV1.CreateCell(18).SetCellValue(Functions.ConvertToString(maestraDescuento[i].DetalleBeneficio));


                    bool parseRestricciones = float.TryParse(maestraDescuento[i].Restricciones, out float Restricciones);
                    if (parseRestricciones) rowDataV1.CreateCell(19).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].Restricciones).Value);
                    else rowDataV1.CreateCell(19).SetCellValue(Functions.ConvertToString(maestraDescuento[i].Restricciones));


                    bool parseDetalleAdicional = float.TryParse(maestraDescuento[i].DetalleAdicional, out float DetalleAdicional);
                    if (parseDetalleAdicional) rowDataV1.CreateCell(20).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].DetalleAdicional).Value);
                    else rowDataV1.CreateCell(20).SetCellValue(Functions.ConvertToString(maestraDescuento[i].DetalleAdicional));


                    bool parseDiasAtencionEstab = float.TryParse(maestraDescuento[i].DiasAtencionEstab, out float DiasAtencionEstab);
                    if (parseDiasAtencionEstab) rowDataV1.CreateCell(21).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].DiasAtencionEstab).Value);
                    else rowDataV1.CreateCell(21).SetCellValue(Functions.ConvertToString(maestraDescuento[i].DiasAtencionEstab));


                    bool parseHorarioEstablec = float.TryParse(maestraDescuento[i].HorarioEstablec, out float HorarioEstablec);
                    if (parseHorarioEstablec) rowDataV1.CreateCell(22).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].HorarioEstablec).Value);
                    else rowDataV1.CreateCell(22).SetCellValue(Functions.ConvertToString(maestraDescuento[i].HorarioEstablec));


                    bool parseRuc = float.TryParse(maestraDescuento[i].Ruc, out float Ruc);
                    if (parseRuc) rowDataV1.CreateCell(23).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].Ruc).Value);
                    else rowDataV1.CreateCell(23).SetCellValue(Functions.ConvertToString(maestraDescuento[i].Ruc));


                    bool parseRazonSocial = float.TryParse(maestraDescuento[i].RazonSocial, out float RazonSocial);
                    if (parseRazonSocial) rowDataV1.CreateCell(24).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].RazonSocial).Value);
                    else rowDataV1.CreateCell(24).SetCellValue(Functions.ConvertToString(maestraDescuento[i].RazonSocial));


                    bool parseDireccion = float.TryParse(maestraDescuento[i].Direccion, out float Direccion);
                    if (parseDireccion) rowDataV1.CreateCell(25).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].Direccion).Value);
                    else rowDataV1.CreateCell(25).SetCellValue(Functions.ConvertToString(maestraDescuento[i].Direccion));


                    bool parseLatitud = float.TryParse(maestraDescuento[i].Latitud, out float Latitud);
                    if (parseLatitud) rowDataV1.CreateCell(26).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].Latitud).Value);
                    else rowDataV1.CreateCell(26).SetCellValue(Functions.ConvertToString(maestraDescuento[i].Latitud));


                    bool parseLongitud = float.TryParse(maestraDescuento[i].Longitud, out float Longitud);
                    if (parseLongitud) rowDataV1.CreateCell(27).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].Longitud).Value);
                    else rowDataV1.CreateCell(27).SetCellValue(Functions.ConvertToString(maestraDescuento[i].Longitud));


                    bool parseFlagOnline = float.TryParse(maestraDescuento[i].FlagOnline, out float FlagOnline);
                    if (parseFlagOnline) rowDataV1.CreateCell(28).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].FlagOnline).Value);
                    else rowDataV1.CreateCell(28).SetCellValue(Functions.ConvertToString(maestraDescuento[i].FlagOnline));


                    bool parseRubroBcpApp = float.TryParse(maestraDescuento[i].RubroBcpApp, out float RubroBcpApp);
                    if (parseRubroBcpApp) rowDataV1.CreateCell(29).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].RubroBcpApp).Value);
                    else rowDataV1.CreateCell(29).SetCellValue(Functions.ConvertToString(maestraDescuento[i].RubroBcpApp));


                    bool parseRubroParaTi = float.TryParse(maestraDescuento[i].RubroParaTi, out float RubroParaTi);
                    if (parseRubroParaTi) rowDataV1.CreateCell(30).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].RubroParaTi).Value);
                    else rowDataV1.CreateCell(30).SetCellValue(Functions.ConvertToString(maestraDescuento[i].RubroParaTi));


                    bool parseRubroParaTiApp = float.TryParse(maestraDescuento[i].RubroParaTiApp, out float RubroParaTiApp);
                    if (parseRubroParaTiApp) rowDataV1.CreateCell(31).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].RubroParaTiApp).Value);
                    else rowDataV1.CreateCell(31).SetCellValue(Functions.ConvertToString(maestraDescuento[i].RubroParaTiApp));


                    bool parseDepartamentoEstablec = float.TryParse(maestraDescuento[i].DepartamentoEstablec, out float DepartamentoEstablec);
                    if (parseDepartamentoEstablec) rowDataV1.CreateCell(32).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].DepartamentoEstablec).Value);
                    else rowDataV1.CreateCell(32).SetCellValue(Functions.ConvertToString(maestraDescuento[i].DepartamentoEstablec));


                    bool parseProvinciaEstablec = float.TryParse(maestraDescuento[i].ProvinciaEstablec, out float ProvinciaEstablec);
                    if (parseProvinciaEstablec) rowDataV1.CreateCell(33).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].ProvinciaEstablec).Value);
                    else rowDataV1.CreateCell(33).SetCellValue(Functions.ConvertToString(maestraDescuento[i].ProvinciaEstablec));


                    bool parseDistritoEstablec = float.TryParse(maestraDescuento[i].DistritoEstablec, out float DistritoEstablec);
                    if (parseDistritoEstablec) rowDataV1.CreateCell(34).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].DistritoEstablec).Value);
                    else rowDataV1.CreateCell(34).SetCellValue(Functions.ConvertToString(maestraDescuento[i].DistritoEstablec));


                    bool parseSegmentoCliente = float.TryParse(maestraDescuento[i].SegmentoCliente, out float SegmentoCliente);
                    if (parseSegmentoCliente) rowDataV1.CreateCell(35).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].SegmentoCliente).Value);
                    else rowDataV1.CreateCell(35).SetCellValue(Functions.ConvertToString(maestraDescuento[i].SegmentoCliente));


                    bool parseUrl = float.TryParse(maestraDescuento[i].Url, out float Url);
                    if (parseUrl) rowDataV1.CreateCell(36).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].Url).Value);
                    else rowDataV1.CreateCell(36).SetCellValue(Functions.ConvertToString(maestraDescuento[i].Url));


                    bool parseSexo = float.TryParse(maestraDescuento[i].Sexo, out float Sexo);
                    if (parseSexo) rowDataV1.CreateCell(37).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].Sexo).Value);
                    else rowDataV1.CreateCell(37).SetCellValue(Functions.ConvertToString(maestraDescuento[i].Sexo));


                    bool parseMensajeMotivador = float.TryParse(maestraDescuento[i].MensajeMotivador, out float MensajeMotivador);
                    if (parseMensajeMotivador) rowDataV1.CreateCell(38).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].MensajeMotivador).Value);
                    else rowDataV1.CreateCell(38).SetCellValue(Functions.ConvertToString(maestraDescuento[i].MensajeMotivador));


                    bool parseTipoSubscripcion = float.TryParse(maestraDescuento[i].TipoSubscripcion, out float TipoSubscripcion);
                    if (parseTipoSubscripcion) rowDataV1.CreateCell(39).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].TipoSubscripcion).Value);
                    else rowDataV1.CreateCell(39).SetCellValue(Functions.ConvertToString(maestraDescuento[i].TipoSubscripcion));


                    bool parseTipoOrigen = float.TryParse(maestraDescuento[i].TipoOrigen, out float TipoOrigen);
                    if (parseTipoOrigen) rowDataV1.CreateCell(40).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].TipoOrigen).Value);
                    else rowDataV1.CreateCell(40).SetCellValue(Functions.ConvertToString(maestraDescuento[i].TipoOrigen));


                    bool parseCodPromocion = float.TryParse(maestraDescuento[i].CodPromocion, out float CodPromocion);
                    if (parseCodPromocion) rowDataV1.CreateCell(41).SetCellValue(Functions.ConvertToFloat(maestraDescuento[i].CodPromocion).Value);
                    else rowDataV1.CreateCell(41).SetCellValue(Functions.ConvertToString(maestraDescuento[i].CodPromocion));


                    bool parseFechaProceso = float.TryParse(maestraDescuento[i].FechaProceso, out float FechaProceso);
                    if (parseFechaProceso) rowDataV1.CreateCell(42).SetCellValue(maestraDescuento[i].FechaProceso);
                    else rowDataV1.CreateCell(42).SetCellValue(maestraDescuento[i].FechaProceso);

                }
                
                string filePath = $"{pathSalida}/{fileName}";
                using (FileStream file = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write))
                {
                    hssfworkbook.Write(file);
                    file.Close();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("El error es el siguiente: {0}", ex.Message);
            }
            hssfworkbook = null;
        }
    }
}
