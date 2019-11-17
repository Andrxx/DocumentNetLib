﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentNetLib.DAL;
using SautinSoft.Document;
using DocumentNetLib.BLL;

namespace DocumentNetLib.Controllers
{
    public class HomeController : Controller
    {
        DocumentManager mng = new DocumentManager();
        public ActionResult Index()
        {
            //string docPath = @"TestTemplate.docx";
            //string templatePath = @"TestTemplate.docx";
            DocumentCore doc = mng.createTemplate();
            mng.SaveDocumentAs(doc, @"TemplateFromMetod.docx");

            //mng.CreateDocument(docPath);
            DocumentCore document = mng.LoadDocument(@"TemplateFromMetod.docx");
            if (document != null)
            {
                User user = new User { FirstName = "John", LastName = "Smit", Guid = "12" };
                DocumentCore updatedDocument = mng.ChangeDocument(document, user);
                string newDocName = "UserTemplate" + user.Guid + ".docx";
                mng.SaveDocumentAs(document, newDocName);
                System.Diagnostics.Process.Start(@"TemplateFromMetod.docx");
                System.Diagnostics.Process.Start(newDocName);
            }

            System.Diagnostics.Process.Start(@"TemplateFromMetod.docx");
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