using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using CsvHelper;
using CsvHelper.Configuration;
using CsvToXlsxWebApp.Services;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Microsoft.Extensions.Logging;

namespace CsvToXlsxWebApp.Controllers
{
    public class CsvController : Controller
    {
        private readonly CsvService _csvService;
        private readonly ILogger<CsvController> _logger;

        public CsvController(CsvService csvService, ILogger<CsvController> logger)
        {
            _csvService = csvService;
            _logger = logger;
        }

        [HttpGet]
        public IActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                ViewBag.ErrorMessage = "Please select a CSV file to upload.";
                return View();
            }

            using (var reader = new StreamReader(file.OpenReadStream()))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                var records = csv.GetRecords<dynamic>().ToList();
                if (records.Any())
                {
                    var headers = ((IDictionary<string, object>)records.First()).Keys.ToList();
                   ViewBag.Headers = headers;

            // Convert records to a more suitable format for the view
            var formattedRecords = records.Select(record =>
                headers.ToDictionary(
                    header => header, 
                    header => ((IDictionary<string, object>)record).ContainsKey(header) ? ((IDictionary<string, object>)record)[header] : null
                )
            ).ToList();

            ViewBag.Records = formattedRecords;
            _csvService.SetRecords(records);
                }
                else
                {
                    ViewBag.ErrorMessage = "No records found in the uploaded file.";
                }
            }

            return View();
        }

        [HttpGet]
        public IActionResult Download()
        {
            var records = _csvService.GetRecords();
            if (records == null || !records.Any())
            {
                _logger.LogWarning("No data available for download.");
                return BadRequest("No data available for download.");
            }

            var fileContent = _csvService.ExportToExcel(records);
            var fileName = "data.xlsx";
            return File(fileContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        [HttpGet]
public IActionResult DownloadData()
{
    var records = _csvService.GetRecords();
    if (records == null || !records.Any())
    {
        _logger.LogWarning("No data available.");
        return Json(new { error = "No data available." });
    }

    var data = records.Select(record => (IDictionary<string, object>)record).ToList();
    return Json(data);
}

    }
}
