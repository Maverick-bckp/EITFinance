using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace EITFinance.Controllers
{
    public class DashboardController : Controller
    {
        IHttpContextAccessor _httpContextAccessor;
        public DashboardController(IHttpContextAccessor httpContextAccessor)
        {
            _httpContextAccessor = httpContextAccessor;
        }
        public IActionResult Index()
        {
            string username = _httpContextAccessor.HttpContext.Session.GetString("username");
            if (username == null)
            {
                return RedirectToAction("Index", "Login");
            }
            return View();
        }
    }
}
