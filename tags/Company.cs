using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace deliverynote2xml.tags
{
    [Serializable]
    public class Company
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string Zip { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Country { get; set; }

        //public string FiscalCode { get; set; }
        //public string VatCode { get; set; }
    }
}
