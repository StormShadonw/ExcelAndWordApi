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
        string sheetName = "Hoja1";
        string connectionString = "DefaultEndpointsProtocol=https;AccountName=riascosservicesstorage;AccountKey=ycHtf5e4s/n4dH4dCnr9ayVro8Ka0nywY9uqwJba50mn1LGfZM6CWI2pTclU3XWnSHom8oc5oc9L+AStqBHiKA==;EndpointSuffix=core.windows.net";
        string containerName = "generalcontainer";

        [HttpPost]
        public async Task<IActionResult> Index(string documentId, [FromBody] Person body)
        {
            try
            {
                var blob = await getBlobClient(connectionString, containerName, documentId);
                Stream memoryStream = blob.DownloadStreamingAsync().Result.Value.Content;
                var package = new ExcelPackage(memoryStream);
                var worksheet = package.Workbook.Worksheets[sheetName]; // Cambia "MiHoja" al nombre de tu hoja
                var indextoInsert = worksheet.Dimension.End.Row + 1;
                worksheet.InsertRow(indextoInsert, 1);
                worksheet.Cells["A" + indextoInsert].Value = body.Name;
                worksheet.Cells["B" + indextoInsert].Value = body.Age;
                package.Save();
                uploadBlob(blob, package.Stream);

                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }

        }

        async void uploadBlob(BlobClient blob, Stream stream)
        {
            stream.Position = 0;
            blob.Upload(stream, overwrite: true);
        }

        [HttpPut("{index}")]
        public async Task<IActionResult> Index(string documentId, int index, [FromBody] Person body)
        {
            var blob = await getBlobClient(connectionString, containerName, documentId);
            Stream memoryStream = blob.DownloadStreamingAsync().Result.Value.Content;
            var package = new ExcelPackage(memoryStream);
            var worksheet = package.Workbook.Worksheets[sheetName]; // Cambia "MiHoja" al nombre de tu hoja
            var rowToUpdate = worksheet.Cells["A" + index + ":B" + index];
            rowToUpdate["A" + index].Value = body.Name;
            rowToUpdate["B" + index].Value = body.Age;
            package.Save();
            uploadBlob(blob, package.Stream);
            return Ok();

        }

        [HttpDelete("{index}")]
        public async Task<IActionResult> Index(string documentId, int index)
        {
            var blob = await getBlobClient(connectionString, containerName, documentId);
            Stream memoryStream = blob.DownloadStreamingAsync().Result.Value.Content;
            var package = new ExcelPackage(memoryStream);
            var worksheet = package.Workbook.Worksheets[sheetName]; // Cambia "MiHoja" al nombre de tu hoja
            worksheet.DeleteRow(index);
            package.Save();
            uploadBlob(blob, package.Stream);
            return Ok();
        }

        async Task<BlobClient> getBlobClient(string connectionString, string containerName, string blobName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
            BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
            BlobClient blobClient = containerClient.GetBlobClient(blobName);
            if (await blobClient.ExistsAsync())
            {
                return blobClient;
            }
            else
            {
                return null;
            }
        }



        [HttpGet]
        public async Task<IActionResult> Index(string documentId)
        {
            try
            {
                var blob = await getBlobClient(connectionString, containerName, documentId);
                Stream memoryStream = blob.DownloadStreamingAsync().Result.Value.Content;
                var data = new List<List<object>>();
                using (var package = new ExcelPackage(memoryStream))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets.FirstOrDefault();

                    if (worksheet != null)
                    {
                        var data1 = worksheet.Rows.Range.Value as object[,];
                        for (var c = 0; c < worksheet.Dimension.End.Row; c++)
                        {
                            List<object> datos = new List<object>();
                            for (var d = 0; d < worksheet.Columns.EndColumn; d++)
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
        public async Task<IActionResult> DownloadFile(string documentId)
        {
            try
            {
                var blob = await getBlobClient(connectionString, containerName, documentId);
                BlobDownloadInfo blobDownloadInfo = await blob.DownloadAsync();
                MemoryStream memoryStream = new MemoryStream();
                await blobDownloadInfo.Content.CopyToAsync(memoryStream);
                byte[] blobBytes = memoryStream.ToArray();
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                return File(blobBytes, contentType, "Data.xlsx");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al leer el archivo Excel: {ex.Message}");
            }

        }
    }
}
