using System;
using System.Collections.Generic;

namespace MerginX.Entities
{
    public class MaestraTotal
    {
        public List<MaestraDescuentos> MaestraDescuentos { get; set; }
        public List<MaestraFuentes> MaestraFuentes { get; set; }
        public List<MaestraEstablecimientos> MaestraEstablecimientos { get; set; }

        public MaestraTotal()
        {
        }
    }
}
