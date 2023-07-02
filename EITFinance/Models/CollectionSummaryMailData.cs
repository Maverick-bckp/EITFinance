using Newtonsoft.Json.Linq;
using System.Collections;
using System.Collections.Generic;

namespace EITFinance.Models
{
    public class CollectionSummaryMailData
    {
        public IEnumerable[] mailTo { get; set; }
        public IEnumerable[] CCTo { get; set; }
        public string clientName { get; set; }
        public JArray collectionData { get; set; }
    }
}
