using EITFinance.Models;
using EITFinance.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.Data;

namespace EITFinance.Controllers
{
    public class BillingController : Controller
    {
        IBillingService _billingService;
        IHttpContextAccessor _httpContextAccessor;
        ILoginService _loginService;
        private IConfiguration _configuration;

        public BillingController(IBillingService billingService, IHttpContextAccessor httpContextAccessor
                               , ILoginService loginService,IConfiguration configuration)
        {
            _billingService = billingService;
            _httpContextAccessor = httpContextAccessor;
            _loginService = loginService;
            _configuration = configuration; 
        }
        public IActionResult Index()
        {
            string username = _httpContextAccessor.HttpContext.Session.GetString("username");
            string moduleBillingSPOC = _configuration.GetValue<string>("ModuleBillingSPOC");
            if (username == null)
            {
                return RedirectToAction("Index", "Login");
            }
            else
            {
                DataTable dt_authorized_status = _loginService.getAuthorizationStatus(username, moduleBillingSPOC);
                if (dt_authorized_status.Rows.Count == 0)
                {
                    return RedirectToAction("Index", "Dashboard");
                }
            }
            return View();
        }

        [HttpPost]
        public JsonResult UploadBillingFile(IFormFile file)
        {
            var billUploadStatus = _billingService.uploadBillingFile(file);

            Billing billing = new Billing();
            billing.status = billUploadStatus;

            return Json(billing);
        }

        [HttpPost]
        public JsonResult uploadMaillingAddresses(IFormFile file)
        {
            var billUploadStatus = _billingService.uploadMaillingAddresses(file);

            Billing billing = new Billing();
            billing.status = billUploadStatus;

            return Json(billing);
        }

        [HttpPost]
        public JsonResult SendBillingMail()
        {
            var sendMailStatus = _billingService.sendBillingMail();

            var billing = new Billing();
            billing.status = sendMailStatus;
            return Json(billing);
        }
    }
}
