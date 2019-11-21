using System;
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
            string docPath = @"TestTemplate.docx";
            //string testDoc = @"NewDoc.docx";
            //string templatePath = @"TestTemplate.docx";
            //DocumentCore document = mng.CreateDocument(testDoc);
            DocumentCore document = mng.LoadDocument(docPath);
            if (document != null)
            {
                User user = new User { FirstName = "John", LastName = "Smit", Guid = "12" };
                Dictionary<string, string> dictionary = new Dictionary<string, string>();
                dictionary.Add("{user}", user.FirstName + " " + user.LastName);
                dictionary.Add("{date}", DateTime.Today.Date.ToString("d"));
                dictionary.Add("{val21}", "21");
                dictionary.Add("{val31}", "31");
                dictionary.Add("{val41}", "41");
                dictionary.Add("{val51}", "51");
                dictionary.Add("{val61}", "61");
                dictionary.Add("{val22}", "22");
                dictionary.Add("{val32}", "32");
                dictionary.Add("{val42}", "42");
                dictionary.Add("{val52}", "52");
                dictionary.Add("{val62}", "62");

                DocumentCore updatedDocument = mng.ChangeStringData(document, dictionary);
                updatedDocument = mng.AddTable(updatedDocument, "{table}", 5, 5);
                string newDocName = "UserTemplate" + user.Guid + ".docx";
                mng.SaveDocumentAs(document, newDocName);
                System.Diagnostics.Process.Start(docPath);
                System.Diagnostics.Process.Start(newDocName);
            }


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