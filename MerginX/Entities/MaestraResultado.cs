using System.Collections.Generic;

namespace MerginX.Entities
{
    public class MaestraResultado
    {
        public List<MaestraDescuentos> DescuentosAceptados { get; set; }
        public List<MaestraDescuentos> DescuentosRechazados { get; set; }
    }
}