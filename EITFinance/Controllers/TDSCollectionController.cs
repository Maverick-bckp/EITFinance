using EITFinance.Models;
using EITFinance.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.Data;

namespace EITFinance.Controllers
{
    public class TDSCollectionController : Controller
    {
        ITDSCollectionService _iTDSCollectionService;
        IHttpContextAccessor _httpContextAccessor;
        ILoginService _loginService;
        private IConfiguration _configuration;

        public TDSCollectionController(ITDSCollectionService iTDSCollectionService, IHttpContextAccessor httpContextAccessor
                               , ILoginService loginService, IConfiguration configuration)
        {
            _iTDSCollectionService = iTDSCollectionService;
            _httpContextAccessor = httpContextAccessor;
            _loginService = loginService;
            _configuration = configuration;
        }
        public IActionResult Index()
        {
            string username = _httpContextAccessor.HttpContext.Session.GetString("username");
            string moduleTDSCollection = _configuration.GetValue<string>("ModuleTDSCollection");
            if (username == null)
            {
                return RedirectToAction("Index", "Login");
            }
            else
            {
                DataTable dt_authorized_status = _loginService.getAuthorizationStatus(username, moduleTDSCollection);
                if (dt_authorized_status.Rows.Count == 0)
                {
                    return RedirectToAction("Index", "Dashboard");
                }
            }
            return View();
        }

        [HttpPost]
        public JsonResult UploadTDSCollectionFile(IFormFile file)
        {
            var sendMailStatus = _iTDSCollectionService.UploadTDSCollectionFile(file);

            var tdsCol = new TDSCollectionLog();
            tdsCol.status = sendMailStatus;
            return Json(tdsCol);
        }

        [HttpPost]
        public JsonResult sendTDSReconciliationMail()
        {
            var sendMailStatus = _iTDSCollectionService.sendTDSCollectionMail();

            var tdsCol = new TDSCollectionLog();
            tdsCol.status = sendMailStatus;
            return Json(tdsCol);
        }

    }
}
