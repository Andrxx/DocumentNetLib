using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SautinSoft.Document;

namespace DocumentNetLib.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            string docPath = @"Result.docx";
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hello World!");
            dc.Save(docPath);
            System.Diagnostics.Process.Start(docPath);
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}