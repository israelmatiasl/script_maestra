using System;
using System.Globalization;
using System.Text.RegularExpressions;

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
