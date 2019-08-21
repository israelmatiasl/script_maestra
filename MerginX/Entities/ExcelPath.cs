using System;
namespace MerginX.Entities
{
    public class ExcelPath
    {
        public string Name { get; set; }
        public string Value { get; set; }

        public ExcelPath(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }
}
