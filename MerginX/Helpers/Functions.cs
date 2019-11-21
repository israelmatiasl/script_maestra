using System;
using System.Globalization;
using System.Text.RegularExpressions;
using NPOI.SS.Formula.Functions;
using Match = System.Text.RegularExpressions.Match;

namespace MerginX.Helpers
{
    public static class Functions
    {
        public static int ConvertToInt(string value)
        {
            int valueConverted;
            var converted = Int32.TryParse(value, out valueConverted);
            if (!converted)
            {
                return 0;
            }

            return valueConverted;
        }

        public static string ConvertToString(string value)
        {
            if(string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value;
        }

        public static float? ConvertToFloat(string value)
        {
            float valueConverted;
            var converted = float.TryParse(value, out valueConverted);
            if (!converted)
            {
                return null;
            }

            return valueConverted;
        }

        public static int? ConvertToInt32(string value)
        {
            int valueConverted;
            var converted = Int32.TryParse(value, out valueConverted);
            if (!converted)
            {
                return null;
            }

            return valueConverted;
        }

        public static int ConvertToDateString(DateTime dateTime)
        {
            var day = dateTime.Day.ToString().PadLeft(2, '0');
            var month = dateTime.Month.ToString().PadLeft(2, '0');
            var year = dateTime.Year.ToString();

            var dateString = year + month + day;
            var dateInt = ConvertToInt(dateString);

            return dateInt;
        }

        public static DateTime ConvertDateFromStringNumeric(string date)
        {
            var year_md = Convert.ToInt32(date.Substring(0, 4));
            var month_md = Convert.ToInt32(date.Substring(4, 2));
            var day_md = Convert.ToInt32(date.Substring(6, 2));
            
            var dateGenerate = new DateTime(year_md, month_md, day_md);
            return dateGenerate;
        }

        public static string ConvertToDateFromRegexDate(string date)
        {
            if (Regex.IsMatch(date, @"(\d{1,2})/(\d{1,2})/(\d{4})"))
            {
                var r = new Regex(@"(\d{1,2})/(\d{1,2})/(\d{4})");
                Match m = r.Match(date);

                try
                {
                    var dateNormal = DateTime.ParseExact(m.Value, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    return ConvertToDateString(dateNormal).ToString();
                }
                catch
                {
                    return ConvertDateValid(m.Value).ToString();
                }
            }

            if (Regex.IsMatch(date, @"(\d{1,2})-(\d{1,2})-(\d{2,4})"))
            {
                var r = new Regex(@"(\d{1,2})-(\d{1,2})-(\d{4})");
                Match m = r.Match(date);

                try
                {
                    var dateNormal = DateTime.ParseExact(m.Value, "dd-MM-yyyy", CultureInfo.InvariantCulture);
                    return ConvertToDateString(dateNormal).ToString();
                }
                catch
                {
                    return ConvertDateValid(m.Value).ToString();
                }
                
            }

            if (Regex.IsMatch(date, @"[0-9][0-9][/][0-9][0-9][/][0-9][0-9]"))
            {
                var r = new Regex(@"[0-9][0-9][/][0-9][0-9][/][0-9][0-9]*");
                Match m = r.Match(date);

                try
                {
                    var dateNormal = DateTime.ParseExact(m.Value, "MM/dd/yy", CultureInfo.InvariantCulture);
                    return ConvertToDateString(dateNormal).ToString();
                }
                catch
                {
                    return ConvertDateValid(m.Value).ToString();
                }
            }

            return null;
        }

        public static int GetIndexNotNumeric(string data)
        {
            int pos = -1;
            for (int i = 0; i < data.Length; i++)
            {
                var converted = Int32.TryParse(data[i].ToString(), out int _);
                if (!converted)
                {
                    pos = i;
                    break;
                }
            }

            return pos;
        }
        
        public static DateTime ConvertFormatDateFromString(string date)
        {
            int pos = GetIndexNotNumeric(date);
            if (pos == -1) throw new InvalidOperationException();

            var dateSplited = date.Split(date[pos]);
                
            var dateFormated = new DateTime(Convert.ToInt32(dateSplited[0]),
                Convert.ToInt32(dateSplited[1]), 
                Convert.ToInt32(dateSplited[2]));

            return dateFormated;
        }
        private static int ConvertDateValid(string date)
        {
            var arrEnteros = date.Split('/');
            if (arrEnteros.Length < 2)
            {
                arrEnteros = date.Split('-');
            }

            if (arrEnteros.Length < 3)
            {
                return 20000000;
            }

            var anio = arrEnteros[2];
            if (anio.Length == 2)
            {
                anio = $"20{arrEnteros[2]}";
            }
            var mes = arrEnteros[1].PadLeft(2, '0');
            var dia = arrEnteros[0].PadLeft(2, '0');

            var fechaCompleta = anio + mes + dia;

            return Convert.ToInt32(fechaCompleta);
        }
    }
}
