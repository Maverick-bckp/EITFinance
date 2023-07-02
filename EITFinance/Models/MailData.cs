using System.Collections;

namespace EITFinance.Models
{
    public class MailData
    {
        public IEnumerable mailTo { get; set; }
        public IEnumerable CCTo { get; set; }
        public IEnumerable mailBody { get; set; }
        public IEnumerable log { get; set; }
    }
}
