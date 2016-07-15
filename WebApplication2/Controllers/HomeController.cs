using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebApplication2.Controllers
{
    public class HomeController : Controller
    {

        public Microsoft.Office.Interop.Word.Document wordDocument { get; set; }

        public ActionResult Index()
        {
            DirectoryInfo salesFTPDirectory = new DirectoryInfo(Server.MapPath("~/App_Data/uploads")); ;
            FileInfo[] files = salesFTPDirectory.GetFiles();
            //files = files.Where(f => f.Extension == ".pdf").OrderBy(f => f.Name);
            var fileslist = files.Where(f => f.Extension == ".pdf").OrderBy(f => f.Name).Select(f => f.Name).ToArray();
            return View(fileslist);
            // return View();
        }



        // This action handles the form POST and the upload
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            // Verify that the user selected a file
            if (file != null && file.ContentLength > 0)
            {
                // extract only the filename
                var fileName = Path.GetFileName(file.FileName);
                // store the file inside ~/App_Data/uploads folder
                var path = Path.Combine(Server.MapPath("~/App_Data/uploads"), fileName);
                var pathpdf = Path.Combine(Server.MapPath("~/App_Data/uploads"), "fileupload_"+DateTime.Now.ToString("yyyyMMdd_hhmm"));
                file.SaveAs(path);



                Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
                wordDocument = appWord.Documents.Open(path);
                wordDocument.ExportAsFixedFormat(pathpdf, WdExportFormat.wdExportFormatPDF);



            }
            // redirect back to the index action to show the form once again

            DirectoryInfo salesFTPDirectory = new DirectoryInfo(Server.MapPath("~/App_Data/uploads")); ;
            FileInfo[] files = salesFTPDirectory.GetFiles();
            //files = files.Where(f => f.Extension == ".pdf").OrderBy(f => f.Name);
            var fileslist = files.Where(f => f.Extension == ".pdf").OrderBy(f => f.Name).Select(f => f.Name).ToArray();
            return View(fileslist);

            //return RedirectToAction("Index");
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