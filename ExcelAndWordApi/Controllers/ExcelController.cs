using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
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

        [HttpPut]
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

        [HttpDelete]
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
        public async Task<IActionResult> DownloadFile()
        {
            try
            {
                string connectionString = "DefaultEndpointsProtocol=https;AccountName=riascosservicesstorage;AccountKey=ycHtf5e4s/n4dH4dCnr9ayVro8Ka0nywY9uqwJba50mn1LGfZM6CWI2pTclU3XWnSHom8oc5oc9L+AStqBHiKA==;EndpointSuffix=core.windows.net"; // Debes reemplazar con tu cadena de conexión real
                string containerName = "generalcontainer";
                string blobName = "Document.docx";

                BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                BlobClient blobClient = containerClient.GetBlobClient(blobName);
                var exits = await blobClient.ExistsAsync();
                if (exits)
                {
                    BlobDownloadInfo blobDownloadInfo = await blobClient.DownloadAsync();
                    MemoryStream memoryStream = new MemoryStream();
                    await blobDownloadInfo.Content.CopyToAsync(memoryStream);
                    byte[] blobBytes = memoryStream.ToArray();
                    var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    return File(blobBytes, contentType, "Data.xlsx");


                }
                else
                {
                    return NotFound("El archivo Excel no existe.");
                }

            }
            catch (Exception ex)
            {
                return BadRequest($"Error al leer el archivo Excel: {ex.Message}");
            }

        }
    }
}
