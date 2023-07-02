using EITFinance.Models;

namespace EITFinance.Services
{
    public interface ISMTPService
    {
        void sendMail(CollectionSummaryMailData mailBody);
    }
}
