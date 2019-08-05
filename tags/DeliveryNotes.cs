using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.Schema;

namespace deliverynote2xml.tags
{
    [Serializable]
    public class DeliveryNotes
    {
        public DeliveryNotes()
        {
            this.Company = new Company();
            this.Documents = new List<Document>();
        }

        public Company Company { get; set; }

        public List<Document> Documents { get; set; }
    }
}
