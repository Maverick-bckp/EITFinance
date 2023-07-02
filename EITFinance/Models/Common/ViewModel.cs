using BundleTransformer.Core.Transformers;
using EITFinance.Models.Timesheet;
using System.Collections.Generic;

namespace EITFinance.Models.Common
{
    public class ViewModel
    {
        public FiscalYear FiscalYear { get; set; }
        public ClientMaster ClientMaster { get; set; }
        public MonthYear monthYear { get; set; }
    }
}
