using EITFinance.Models;
using EITFinance.Models.Timesheet;
using EITFinance.Models.Timesheet.DTOs;
using EITFinance.Services;
using EITFinance.Utilities;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace EITFinance.Controllers
{
    public class TimesheetController : Controller
    {
        private ITimesheetService _timesheetService;
        private IEmailSender _emailSender;
        private readonly IConfiguration _configuration;
        private readonly ILogger<TimesheetController> _logger;
        IHttpContextAccessor _httpContextAccessor;
        public TimesheetController(ITimesheetService timesheetService, IEmailSender emailSender, IConfiguration configuration, ILogger<TimesheetController> logger, IHttpContextAccessor httpContextAccessor)
        {
            _timesheetService = timesheetService;
            _emailSender = emailSender;
            _configuration = configuration;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
        }

        public IActionResult Index()
        {
            /*------ Binding Data Into Dropdown ------*/
            ViewBag.ListOfFiscalYear = _timesheetService.getFiscalYearDDLList();
            ViewBag.ListOfClientMaster = _timesheetService.getClientMasterDDLList();
            ViewBag.ListOfMonthYear = _timesheetService.getMonthYearDDLList();

            return View();
        }

        public IActionResult processTimesheet()
        {
            _logger.LogInformation("Timesheet process has been started.");

            _timesheetService.TimesheetProcessor();
            _logger.LogInformation("Timesheet process has been completed.");

            return Json(new { Message = "Process has been completed !" });
        }

        [HttpPost]
        [RequestSizeLimit(10000000)]
        public JsonResult UploadInvoiceFile(IFormFile file)
        {
            var timesheet = new Timesheet();

            string folderPath = _configuration.GetValue<string>("Application:folderPath");

            if (folderPath == null)
            {
                _logger.LogInformation("Application folder paths are missing Application:folderPath or Application:archiveFolderPath.");
                timesheet.status = false;
                return Json(timesheet);
            }

            var status = _timesheetService.uploadInvoiceFile(file, folderPath);
            timesheet.status = status;


            return Json(timesheet);
        }



    }
}
