using EITFinance.Models.Common;

namespace EITFinance.Models.Timesheet
{
    public class Timesheet : BaseEntity
    {
        public string ClientName { get; set; }
        public string filePath { get; set; }
        public bool status { get; set; }

    }
}
