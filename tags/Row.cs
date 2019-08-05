using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace deliverynote2xml.tags
{
    [Serializable]
    public class Row
    {
        public string Code { get; set; }
        public string Description { get; set; }
        public string Qty { get; set; }
        public string Discounts { get; set; }
    }
}
