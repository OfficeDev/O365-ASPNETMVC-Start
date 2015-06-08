using System.Web.Mvc;
using System.Configuration;
using System;

namespace O365_APIs_Start_ASPNET_MVC.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            if (ConfigurationManager.AppSettings["ida:ClientID"] == null)
            {
                ViewBag.DidNotAddConnectedServices = true;
            }
            else
            {
                ViewBag.DidNotAddConnectedServices = false;
            }
            
            bool isEmpty = String.IsNullOrEmpty(ConfigurationManager.AppSettings["ida:TenantID"]);

            if (isEmpty)
            {
                ViewBag.TenantIDIsNull = true;
            }
            else
                ViewBag.TenantIDIsNull = false;

            return View();
        }
    }
}