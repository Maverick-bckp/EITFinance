using Microsoft.AspNetCore.Http;

namespace EITFinance.Services
{
    public interface IBillingService
    {
        bool uploadBillingFile(IFormFile file);
        bool sendBillingMail();
        int deleteBillingSPOCTable(string loginID);
        dynamic getBillingSPOCLog(string loginId);
        bool uploadMaillingAddresses(IFormFile file);
    }
}
