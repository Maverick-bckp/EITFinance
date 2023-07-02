using EITFinance.Models;
using EITFinance.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.Data;

namespace EITFinance.Controllers
{
    public class MaillingAddressController : Controller
    {
        IMaillingAddressService _maillingAddressService;
        IHttpContextAccessor _httpContextAccessor;
        ILoginService _loginService;
        private IConfiguration _configuration;

        public MaillingAddressController(IMaillingAddressService maillingAddressService, IHttpContextAccessor httpContextAccessor
                               , ILoginService loginService, IConfiguration configuration)
        {
            _maillingAddressService = maillingAddressService;
            _httpContextAccessor = httpContextAccessor;
            _loginService = loginService;
            _configuration = configuration;
        }
        public IActionResult Index()
        {
            string username = _httpContextAccessor.HttpContext.Session.GetString("username");
            string moduleMaiilingAddress = _configuration.GetValue<string>("ModuleMaillingAddress");
            if (username == null)
            {
                return RedirectToAction("Index", "Login");
            }
            else
            {
                DataTable dt_authorized_status = _loginService.getAuthorizationStatus(username, "mailling_address");
                if (dt_authorized_status.Rows.Count == 0)
                {
                    return RedirectToAction("Index", "Dashboard");
                }
            }
            return View();
        }

        [HttpPost]
        public JsonResult uploadMaillingAddresses(IFormFile file)
        {
            var billUploadStatus = _maillingAddressService.uploadMaillingAddresses(file);

            Billing billing = new Billing();
            billing.status = billUploadStatus;

            return Json(billing);
        }
    }
}
