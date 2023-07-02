using EITFinance.Services;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using MimeKit;
using Newtonsoft.Json.Linq;
using System;
using System.Globalization;
using System.IO;
using System.Linq;

namespace EITFinance.Controllers
{
    public class POController : Controller
    {
        private IPOService _poService;
        IConfiguration _configuration;
        public POController(IPOService poService, IConfiguration configuration)
        {
            _poService = poService;
            _configuration = configuration;
        }
        public IActionResult Index()
        {
            var host = _configuration.GetValue<string>("MailSettings:Host");
            var port = _configuration.GetValue<int>("MailSettings:Port");
            var username = _configuration.GetValue<string>("MailSettings:UserName");
            var password = _configuration.GetValue<string>("MailSettings:Password");
            var Sender = _configuration.GetValue<string>("MailSettings:Sender");

            /*--- Sender ----*/
            var email = new MimeMessage();
            email.From.Add(MailboxAddress.Parse(Sender));

            email.To.Add(MailboxAddress.Parse("sunil.kumar2@manpowergroup.com"));
          
            var body = new BodyBuilder();
            email.Subject = $"Remittance Advice Required - Client A";
            string htmlBody = @"<p style=""font-family:Calibri;font-size:15;"">Dear Sir/Madam,</p></n><p style=""font-family:Calibri;font-size:15;"">Kindly help us with the invoice-wise adjustment details (Remittance Advice) for the below payments urgently.</p></br>";

            body.HtmlBody = string.Format(htmlBody);
            email.Body = body.ToMessageBody();

            /*--- Connect SMTP server and send mail ----*/
            using var smtp = new SmtpClient();
            smtp.Connect(host, port, false);
            smtp.Authenticate("ayan.roy@in.experis.com", "zzzvvvvqqqq4M+m");
            smtp.Send(email);
            smtp.Disconnect(true);
            return View();
        }


        [HttpPost]
        public void UploadMailingList(IFormFile file)
        {
            var mailUploadStatus = _poService.UploadMaillingAddresses(file);

            //UnbilledRevenue unbilledRevenue = new UnbilledRevenue();
            //unbilledRevenue.status = mailUploadStatus;

            //return Json();
        }
    }
}
