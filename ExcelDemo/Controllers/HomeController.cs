using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using System.IO;

namespace ExcelDemo.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult GetExeclFile()
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Report");
            Sheet.Cells["A1"].Value = "Name";
            Sheet.Cells["B1"].Value = "Department";
            Sheet.Cells["C1"].Value = "Address";
            Sheet.Cells["D1"].Value = "City";
            Sheet.Cells["E1"].Value = "Country";
            Sheet.Cells["A:AZ"].AutoFitColumns();

            //Response.Clear();
            //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //Response.AddHeader("content-disposition", "attachment: filename=" + "Report.xlsx");
            //Response.BinaryWrite(Ep.GetAsByteArray());
            //Response.End();

            return File(Ep.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "File.xlsx");
        }
    }
}