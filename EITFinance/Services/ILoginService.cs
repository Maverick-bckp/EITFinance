namespace EITFinance.Services
{
    public interface ILoginService
    {
        bool authenticate(string username, string password);
        bool Logout();
        dynamic getAuthorizationStatus(string loginID, string module);
    }
}
