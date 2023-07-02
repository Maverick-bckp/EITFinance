using EITFinance.Models;
using EITFinance.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using System;

namespace EITFinance.Controllers
{
    public class LoginController : Controller
    {
        ILoginService _login;
        public LoginController(ILoginService login)
        {
            _login = login;
        }
        public IActionResult Index()
        {
            ViewBag.successStatus = TempData["successStatus"];
            return View();
        }

        public IActionResult Authenticate(Login login)
        {
            if (string.IsNullOrEmpty(login.Username) || string.IsNullOrEmpty(login.Password))
            {
                TempData["successStatus"] = false;
                return RedirectToAction("Index");
            }
            else
            {
                bool status = _login.authenticate(login.Username, login.Password);
                if (status)
                {
                    return RedirectToAction("Index", "Dashboard");
                }
                else
                {
                    TempData["successStatus"] = false;
                    return RedirectToAction("Index");
                }
            }
        }

        public IActionResult Logout()
        {
            bool logoutStatus = _login.Logout();

            return RedirectToAction("Index");
        }
    }
}
