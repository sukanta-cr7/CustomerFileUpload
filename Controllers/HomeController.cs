using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DemoMVCProject1.Models;
using DemoMVCProject1.Services;

namespace DemoMVCProject1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                var path = Path.Combine(Server.MapPath("~/App_Data/Uploads"), Path.GetFileName(file.FileName));
                file.SaveAs(path);
                TempData["UploadedFilePath"] = path;
                TempData["Filename"] = Path.GetFileName(file.FileName);
                return RedirectToAction("MapColumns");
            }

            ViewBag.Error = "File upload failed.";
            return View("Index");
        }

        public ActionResult MapColumns()
        {
            string path = TempData["UploadedFilePath"] as string;
            if (string.IsNullOrEmpty(path))
                return RedirectToAction("Index");

            var headers = ExcelService.GetExcelHeaders(path);
            ViewBag.Headers = headers;

            var dbColumns = typeof(Customer).GetProperties().Select(p => p.Name).ToList();
            ViewBag.DbColumns = dbColumns;

            return View();
        }
    }
}