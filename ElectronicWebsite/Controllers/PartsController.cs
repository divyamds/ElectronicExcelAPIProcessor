using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;

namespace ElectronicWebsite.Controllers
{
    public class PartsController : Controller
    {
        private readonly HttpClient _http;

        public PartsController(IHttpClientFactory httpClientFactory)
        {
            _http = httpClientFactory.CreateClient();
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public async Task<IActionResult> ProcessExcel(IFormFile excelFile, string apiUrl)
        {
            if (excelFile == null || excelFile.Length == 0 || string.IsNullOrEmpty(apiUrl))
            {
                ViewBag.Error = "Please upload an Excel file and enter API URL.";
                return View("Index");
            }

            // ✅ Validate file extension
            var extension = Path.GetExtension(excelFile.FileName).ToLowerInvariant();
            var allowedExtensions = new[] { ".xlsx", ".xlsm", ".xltx", ".xltm" };
            if (!allowedExtensions.Contains(extension))
            {
                ViewBag.Error = "Invalid file format. Please upload a valid Excel file (.xlsx, .xlsm, .xltx, .xltm).";
                return View("Index");
            }

            // ✅ Save with correct extension
            var tempFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}{extension}");
            using (var stream = new FileStream(tempFile, FileMode.Create))
            {
                await excelFile.CopyToAsync(stream);
            }

            using (var workbook = new XLWorkbook(tempFile))
            {
                var ws = workbook.Worksheet(1);
                var headerRow = ws.Row(1);

                // ✅ Map column names dynamically
                var colMap = new Dictionary<string, int>();
                foreach (var cell in headerRow.CellsUsed())
                {
                    var colName = cell.GetString().Trim();
                    if (!string.IsNullOrEmpty(colName))
                        colMap[colName] = cell.Address.ColumnNumber;
                }

                if (!colMap.ContainsKey("PartNumber"))
                {
                    ViewBag.Error = "Excel must contain a 'PartNumber' column.";
                    return View("Index");
                }

                var rows = ws.RangeUsed().RowsUsed().Skip(1); // skip header

                foreach (var row in rows)
                {
                    var partNumber = row.Cell(colMap["PartNumber"]).GetString();
                    if (string.IsNullOrWhiteSpace(partNumber)) continue;

                    var url = apiUrl.Replace("{part_number}", partNumber);
                    try
                    {
                        var response = await _http.GetStringAsync(url);
                        var json = JObject.Parse(response);
                        var firstResult = json["results"]?.FirstOrDefault();
                        if (firstResult == null) continue;

                        foreach (var kv in colMap)
                        {
                            string value = null;

                            // Always keep PartNumber as is
                            if (kv.Key == "PartNumber")
                            {
                                value = row.Cell(kv.Value).GetString(); // keep original
                            }
                            else
                            {
                                // Fill only API columns
                                switch (kv.Key)
                                {
                                    case "Manufacturer":
                                        value = firstResult["manufacturer"]?.ToString();
                                        break;
                                    case "Description":
                                        value = firstResult["description"]?.ToString();
                                        break;
                                    case "Lifecycle":
                                        value = firstResult["lifecycle"]?.ToString();
                                        break;
                                    case "Price":
                                        value = firstResult["price"]?["USD"]?.ToString();
                                        break;
                                    case "Stock":
                                        value = firstResult["stock"]?.ToString();
                                        break;
                                    case "RepresentativeParts":
                                        var reps = firstResult["representativeParts"]?
                                            .Select(r => r["partNumber"]?.ToString());
                                        value = string.Join(", ", reps ?? new List<string>());
                                        break;
                                }

                                // Highlight only missing API values
                                if (string.IsNullOrWhiteSpace(value))
                                {
                                    row.Cell(kv.Value).Style.Fill.BackgroundColor = XLColor.LightPink;
                                }
                            }

                            // Write value to cell
                            row.Cell(kv.Value).Value = value;
                        }


                    }
                    catch
                    {
                        // Skip errors gracefully
                        continue;
                    }
                }

                // ✅ Formatting before saving
                var headerRowStyle = ws.Row(1).Style;
                headerRowStyle.Font.Bold = true;
                headerRowStyle.Fill.BackgroundColor = XLColor.LightGray;
                headerRowStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Columns().AdjustToContents();

                var usedRange = ws.RangeUsed();
                usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // ✅ Save processed file
                var outputFile = Path.Combine(Path.GetTempPath(), $"Processed_{Guid.NewGuid()}.xlsx");
                workbook.SaveAs(outputFile);
                TempData["ProcessedFile"] = outputFile;
            }

            return RedirectToAction("Download");
        }



        [HttpGet]
        public IActionResult Download()
        {
            var filePath = TempData["ProcessedFile"] as string;
            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                return NotFound("No processed file found.");

            var fileName = Path.GetFileName(filePath);

            // Pass info to the view
            ViewBag.FileName = fileName;
            ViewBag.FilePath = filePath;

            return View();
        }

        [HttpGet]
        public IActionResult DownloadFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                return NotFound("File not found.");

            var fileName = Path.GetFileName(filePath);
            var bytes = System.IO.File.ReadAllBytes(filePath);

            return File(bytes,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName);
        }

    }
}
