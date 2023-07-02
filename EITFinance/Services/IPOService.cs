using Microsoft.AspNetCore.Http;

namespace EITFinance.Services
{
    public interface IPOService 
    {
        bool UploadMaillingAddresses(IFormFile file);

        void ProcessPO();
    }
}
