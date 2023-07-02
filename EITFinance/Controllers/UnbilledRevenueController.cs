using EITFinance.Models;
using EITFinance.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.Data;

namespace EITFinance.Controllers
{
    public class UnbilledRevenueController : Controller
    {
        IUnbilledRevenueService _unbilledRevenueService;
        IHttpContextAccessor _httpContextAccessor;
        ILoginService _loginService;
        private IConfiguration _configuration;

        public UnbilledRevenueController(IUnbilledRevenueService unbilledRevenueService, IHttpContextAccessor httpContextAccessor
                               , ILoginService loginService, IConfiguration configuration)
        {
            _unbilledRevenueService = unbilledRevenueService;
            _httpContextAccessor = httpContextAccessor;
            _loginService = loginService;
            _configuration = configuration;
        }


        public IActionResult Index()
        {
            string username = _httpContextAccessor.HttpContext.Session.GetString("username");
            string moduleUnbilledRevenue = "unbilled_revenue";
            if (username == null)
            {
                return RedirectToAction("Index", "Login");
            }
            else
            {
                DataTable dt_authorized_status = _loginService.getAuthorizationStatus(username, moduleUnbilledRevenue);
                if (dt_authorized_status.Rows.Count == 0)
                {
                    return RedirectToAction("Index", "Dashboard");
                }
            }
            return View();
        }


        [HttpPost]
        public JsonResult UploadMailingList(IFormFile file)
        {
            var mailUploadStatus = _unbilledRevenueService.uploadMaillingAddress(file);

            UnbilledRevenue unbilledRevenue = new UnbilledRevenue();
            unbilledRevenue.status = mailUploadStatus;

            return Json(unbilledRevenue);
        }

        [HttpPost]
        public JsonResult UploadUnbilledRevenueFile(IFormFile file)
        {
            var billUploadStatus = _unbilledRevenueService.uploadUnbilledRevenueFile(file);

            UnbilledRevenue unbilledRevenue = new UnbilledRevenue();
            unbilledRevenue.status = billUploadStatus;

            return Json(unbilledRevenue);
        }

        [HttpPost]
        public JsonResult sendUnbilledRevenueDetailsMail()
        {
            var sendMailStatus = _unbilledRevenueService.sendUnbilledRevenueDetailsMail();

            UnbilledRevenue unbilledRevenue = new UnbilledRevenue();
            unbilledRevenue.status = sendMailStatus;

            return Json(unbilledRevenue);
        }

    }
}
