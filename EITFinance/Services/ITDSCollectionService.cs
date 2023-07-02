using Microsoft.AspNetCore.Http;

namespace EITFinance.Services
{
    public interface ITDSCollectionService
    {
        bool UploadTDSCollectionFile(IFormFile file);
        int deleteFromTDSCollectionLogTable(string loginID);
        bool sendTDSCollectionMail();
        int updateMailSendStatus(string clientName);
        dynamic getTDSCollectionLog(string loginId);
    }
}
