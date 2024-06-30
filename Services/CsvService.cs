using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml; 
using Microsoft.Extensions.Logging;

namespace CsvToXlsxWebApp.Services
{
    public class CsvService
    {
        private List<dynamic> _records = new List<dynamic>(); // Temporary storage for records
        private readonly ILogger<CsvService> _logger;

        public CsvService(ILogger<CsvService> logger)
        {
            _logger = logger;
        }

        public IEnumerable<dynamic> GetRecords()
        {
            _logger.LogInformation("Getting records, count: {Count}", _records.Count);
            return _records;
        }

        public void SetRecords(IEnumerable<dynamic> records)
        {
            _records = records.ToList();
            _logger.LogInformation("Setting records, count: {Count}", _records.Count);
        }

        public byte[] ExportToExcel(IEnumerable<dynamic> records)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                if (records != null && records.Any())
                {
                    var headers = ((IDictionary<string, object>)records.First()).Keys.ToList();
                    for (int i = 0; i < headers.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = headers[i];
                    }

                    int row = 2;
                    foreach (var record in records)
                    {
                        int col = 1;
                        foreach (var value in ((IDictionary<string, object>)record).Values)
                        {
                            worksheet.Cells[row, col++].Value = value;
                        }
                        row++;
                    }
                }

                return package.GetAsByteArray();
            }
        }
    }
}
