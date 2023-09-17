using ExcelAndWordApi.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.Internal;
using OfficeOpenXml;

namespace ExcelAndWordApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        string filePath = Path.Combine("C:", "Excel");
        //string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Excel");
        string fileName = "Data.xlsx";
        string sheetName = "Hoja1";
        [HttpPost("Create-file")]
        public IActionResult CreateFile()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                string fullPath = Path.Combine(filePath, fileName);
                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }
                using (var package = new ExcelPackage())
                {
                    package.Workbook.Worksheets.Add(sheetName);
                    // Rellenar el archivo Excel como se hizo en el ejemplo anterior

                    // Guardar el paquete en la ubicación deseada
                    System.IO.File.WriteAllBytes(fullPath, package.GetAsByteArray());
                }
                return Ok(new { FilePath = fullPath });
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }

        [HttpPost]
        public IActionResult Index([FromBody] Person body)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(new FileInfo(Path.Combine(filePath, fileName)));
            var worksheet = package.Workbook.Worksheets[sheetName]; // Cambia "MiHoja" al nombre de tu hoja
            var indextoInsert = worksheet.Dimension.End.Row + 1;
            worksheet.InsertRow(indextoInsert, 1);
            worksheet.Cells["A" + indextoInsert].Value = body.Name;
            worksheet.Cells["B" + indextoInsert].Value = body.Age;
            package.Save();
            return Ok();
        }

        [HttpPut("{index}")]
        public IActionResult Index(int index, [FromBody] Person body)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var package = new ExcelPackage(new FileInfo(Path.Combine(filePath, fileName)));
            var worksheet = package.Workbook.Worksheets[sheetName]; // Cambia "MiHoja" al nombre de tu hoja
            var rowToUpdate = worksheet.Cells["A" + index + ":B" + index];
            rowToUpdate["A" + index].Value = body.Name;
            rowToUpdate["B" + index].Value = body.Age;
            package.Save();
            return Ok();
        }

        [HttpDelete("{index}")]
        public IActionResult Index(int index)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var package = new ExcelPackage(new FileInfo(Path.Combine(filePath, fileName)));
            var worksheet = package.Workbook.Worksheets[sheetName]; // Cambia "MiHoja" al nombre de tu hoja
            worksheet.DeleteRow(index);
            package.Save();
            return Ok();
        }

        [HttpGet]
        public IActionResult Index()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var filePath = Path.Combine(this.filePath, this.fileName); // Reemplaza con la ruta de tu archivo Excel
                FileInfo fileInfo = new FileInfo(filePath);
                var data = new List<List<object>>();

                using (var package = new ExcelPackage(fileInfo))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets.FirstOrDefault();
                    
                    if (worksheet != null)
                    {
                        var data1 = worksheet.Rows.Range.Value as object[,];
                        for (var c = 0; c < worksheet.Dimension.End.Row; c++) {
                            List<object> datos = new List<object>();
                            for(var d = 0; d < worksheet.Columns.EndColumn; d++)
                            {
                                datos.Add(data1[c, d]);
                            }
                            data.Add(datos);
                        }
                        return Ok(data);
                    }
                }

                return NotFound("No se encontró la hoja de Excel o no se leyeron datos.");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al leer el archivo Excel: {ex.Message}");
            }

        }

        [HttpGet("download-file")]
        public IActionResult DownloadFile()
        {
            try
            {
                var rutaArchivo = Path.Combine(filePath, fileName);
                if (!System.IO.File.Exists(rutaArchivo))
                {
                    return NotFound("El archivo Excel no existe.");
                }
                var bytes = System.IO.File.ReadAllBytes(rutaArchivo);
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                return File(bytes, contentType, "Data.xlsx");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al leer el archivo Excel: {ex.Message}");
            }

        }
    }
}
