namespace EITFinance.Services
{
    public interface IActiveDirectoryService
    {
        dynamic Authenticate(string username, string password);
        dynamic getUsernameDetails(string username);
    }
}
