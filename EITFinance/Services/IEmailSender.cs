using EITFinance.Models;

namespace EITFinance.Services
{
    public interface IEmailSender
    {
        void SendEmail(Message message);
    }
}
