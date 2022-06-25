using CreateExcelAndPdfFile.Excel.Models;
using FastMember;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CreateExcelAndPdfFile.Excel.Controllers
{
    // Excel için  => Manage Nuget Package dan,  ""EPPlus""  kütüphanesini yükle
    // PDF   için  => Manage Nuget Package dan,  ""ITextSharp.LGPLv2.Core"" ve ""FastMember""  kütüphanelerini yükle
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult GetPdf()
        {
            DataTable dataTable = new DataTable();

            dataTable.Load(ObjectReader.Create(new List<Customer> {
               new Customer{Id = 1,Name = "Yavuz",Surname = "Akşit",Phone = "5553656545",Age=24,Address="Bilecik"},
               new Customer{Id = 2,Name = "Yavuz1",Surname = "Akşit1",Phone = "5553656545",Age=24,Address="Bilecik1"},
               new Customer{Id = 3,Name = "Yavuz2",Surname = "Akşit2",Phone = "5553656545",Age=24,Address="Bilecik2"},
               new Customer{Id = 4,Name = "Yavuz3",Surname = "Akşit3",Phone = "5553656545",Age=24,Address="Bilecik3"},
               new Customer{Id = 5,Name = "Yavuz4",Surname = "Akşit4",Phone = "5553656545",Age=24,Address="Bilecik4"},
            }));

            string filename = Guid.NewGuid() + ".pdf";
            string path = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/documents/"+filename);
            var stream = new FileStream(path, FileMode.Create);

            Document document = new Document(PageSize.A4,25f, 25f, 25f, 25f);
            PdfWriter.GetInstance(document, stream);

            document.Open();
             

            PdfPTable pdfPTable = new PdfPTable(dataTable.Columns.Count);

            //pdfPTable.AddCell("Ad");
            //pdfPTable.AddCell("Soyad");
            //pdfPTable.AddCell("Yas");

            //pdfPTable.AddCell("Emre");
            //pdfPTable.AddCell("Aktürk");
            //pdfPTable.AddCell("27");

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                pdfPTable.AddCell(dataTable.Columns[i].ColumnName);       
            }

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    pdfPTable.AddCell(dataTable.Rows[i][j].ToString());       
                }
            }

            document.Add(pdfPTable);

            document.Close();


            return File("/documents/"+filename,"application/pdf",filename);
        }

        public IActionResult GetExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();

            var excelBlank = excelPackage.Workbook.Worksheets.Add("Calisma1");

            //excelBlank.Cells[1, 1].Value = "Ad";
            //excelBlank.Cells[1, 2].Value = "Soyad";

            //excelBlank.Cells[2, 1].Value = "Emre";
            //excelBlank.Cells[2, 2].Value = "Aktürk";

            excelBlank.Cells["A1"].LoadFromCollection(new List<Customer> {
               new Customer{Id = 1,Name = "Yavuz",Surname = "Akşit",Phone = "5553656545",Age=24,Address="Bilecik"},
               new Customer{Id = 2,Name = "Yavuz1",Surname = "Akşit1",Phone = "5553656545",Age=24,Address="Bilecik1"},
               new Customer{Id = 3,Name = "Yavuz2",Surname = "Akşit2",Phone = "5553656545",Age=24,Address="Bilecik2"},
               new Customer{Id = 4,Name = "Yavuz3",Surname = "Akşit3",Phone = "5553656545",Age=24,Address="Bilecik3"},
               new Customer{Id = 5,Name = "Yavuz4",Surname = "Akşit4",Phone = "5553656545",Age=24,Address="Bilecik4"},
            },true,OfficeOpenXml.Table.TableStyles.Light15);

            var bytes = excelPackage.GetAsByteArray();
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Guid.NewGuid() + ".xlsx");


        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }

    public class Customer
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Surname { get; set; }
        public int Age { get; set; }
        public string Address { get; set; }
        public string Phone { get; set; }
    }
}
