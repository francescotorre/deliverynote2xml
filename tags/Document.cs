using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace deliverynote2xml.tags
{
    [Serializable]
    public class Document
    {
        public Document(){
            this.Rows = new List<Row>();
        }

        public string CustomerCode { get; set; }
        public string SoldToName { get; set; }
        public string ShipToName { get; set; }
        public string ShipToAddress { get; set; }
        public string ShipToZip { get; set; }
        public string ShipToCity { get; set; }
        public string ShipToState { get; set; }
        public string DocumentType { get; set; }
        public string Date { get; set; }
        public string Number { get; set; }

        //public string Numbering { get; set; }  
        //public string TransportReason { get; set; }
        //public string TransportDateTime { get; set; }

        public List<Row> Rows { get; set; }
    }
}
