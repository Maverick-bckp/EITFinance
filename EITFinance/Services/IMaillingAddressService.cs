using Microsoft.AspNetCore.Http;

namespace EITFinance.Services
{
    public interface IMaillingAddressService
    {
        bool uploadMaillingAddresses(IFormFile file);
        int truncateMaillingAddressTable();
    }
}
