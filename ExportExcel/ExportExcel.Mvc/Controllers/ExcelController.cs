using ExportExcel.Mvc.Code;
using ExportExcel.Mvc.Models;
using ExportExcel.Mvc.ViewModels;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Web.Mvc;

//EPPlus
namespace ExportExcel.Mvc.Controllers
{
    public class ExcelController : Controller
    {
        // GET: Excel
        [HttpGet]
        public ActionResult Index()
        {
            TechnologyViewModel model = new TechnologyViewModel();

            return View(model);
        }

        [HttpGet]
        public FileContentResult ExportToExcel()
        {
            List<Technology> technologies = StaticData.Technologies;
            string[] columns = { "Name", "Project", "Developer" };
            byte[] filecontent = ExcelExportHelper.ExportExcel(technologies, "Technology", true, columns);
            return File(filecontent, ExcelExportHelper.ExcelContentType, "Teste.xlsx");
        }


        [HttpGet]
        public ActionResult SafeDownload()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Download()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Test");
                ExcelWorksheet ws = package.Workbook.Worksheets[1];
                ws.Name = "Test"; //Setting Sheet's name
                ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
                ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

                //Merging cells and create a center heading for out table
                ws.Cells[1, 1].Value = "Sample DataTable Export"; // Heading Name
                ws.Cells[1, 1, 1, 10].Merge = true; //Merge columns start and end range
                ws.Cells[1, 1, 1, 10].Style.Font.Bold = true; //Font should be bold
                ws.Cells[1, 1, 1, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Aligmnet is center

                for (var i = 1; i < 11; i++)
                {
                    for (var j = 2; j < 45; j++)
                    {
                        var cell = ws.Cells[j, i];

                        //Setting Value in cell
                        cell.Value = i * (j - 1);
                    }
                }

                var chart = ws.Drawings.AddChart("chart1", eChartType.AreaStacked);
                //Set position and size
                chart.SetPosition(0, 630);
                chart.SetSize(800, 600);

                // Add the data series. 
                var series = chart.Series.Add(ws.Cells["A2:A46"], ws.Cells["B2:B46"]);

                var memoryStream = package.GetAsByteArray();
                var fileName = string.Format("MyData-{0:yyyy-MM-dd-HH-mm-ss}.xlsx", DateTime.UtcNow);
                // mimetype from http://stackoverflow.com/questions/4212861/what-is-a-correct-mime-type-for-docx-pptx-etc
                return base.File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }
    }
}