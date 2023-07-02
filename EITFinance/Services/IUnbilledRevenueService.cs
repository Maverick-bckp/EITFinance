using Microsoft.AspNetCore.Http;
using System;

namespace EITFinance.Services
{
    public interface IUnbilledRevenueService
    {
        bool uploadUnbilledRevenueFile(IFormFile file);
        int deleteFromUnbilledRevenueTable(string loginID);
        dynamic getUnbilledDataByActions(int actionType = 0, string loginId = null, string parentCategory = null, string clientName = null);
        void sendMail(string[] mailTo, string[] CCTo, string htmlBody, string subject, string attachmentPath, string loginID);
        bool sendUnbilledRevenueDetailsMail();
        bool uploadMaillingAddress(IFormFile file);
        int deleteFromUnbilledRevenueMaillingAddressStagingTable(string loginID);
    }
}
