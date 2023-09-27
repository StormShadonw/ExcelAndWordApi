using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using ExcelAndWordApi.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.Internal;
using OfficeOpenXml;
using System.Linq;

namespace ExcelAndWordApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        string connectionString = "DefaultEndpointsProtocol=https;AccountName=riascosservicesstorage;AccountKey=ycHtf5e4s/n4dH4dCnr9ayVro8Ka0nywY9uqwJba50mn1LGfZM6CWI2pTclU3XWnSHom8oc5oc9L+AStqBHiKA==;EndpointSuffix=core.windows.net";


        [HttpPost("{containerName}/{documentName}")]
        public async Task<IActionResult> Index([FromHeader(Name = "ApiKey")] string apikey, string containerName, string documentName, [FromBody] FT_FD_2558 body)
        {
            try
            {
                containerName = containerName.ToLower();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                BlobClient blob = containerClient.GetBlobClient(documentName);
                
                Stream memoryStream = blob.DownloadStreamingAsync().Result.Value.Content;
                var data = new List<List<object>>();
                using (var package = new ExcelPackage(memoryStream))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets.FirstOrDefault();
                    worksheet.Cells["P8"].Value = body.RazonSocial;
                    worksheet.Cells["CD6"].Value = body.Expediente;
                    worksheet.Cells["P9"].Value = body.TipoIdentificacion;
                    worksheet.Cells["AV9"].Value = body.NoIdentificacion;
                    worksheet.Cells["P8"].Value = body.RazonSocial;
                    for(var c = 0; c < body.Items.Count; c++)
                    {
                        var index = 14 + c;
                        worksheet.Cells["A" + index].Value = c + 1;
                        worksheet.Cells["G" + index].Value = body.Items[c].Descripción;
                        worksheet.Cells["AW" + index].Value = body.Items[c].FechaAño;
                        worksheet.Cells["BA" + index].Value = body.Items[c].FechaMes;
                        worksheet.Cells["BD" + index].Value = body.Items[c].FechaDía;
                        worksheet.Cells["BG" + index].Value = body.Items[c].FolioFI;
                        worksheet.Cells["BJ" + index].Value = body.Items[c].FolioFF;
                        worksheet.Cells["BM" + index].Value = body.Items[c].FolioST;
                        worksheet.Cells["BP" + index].Value = body.Items[c].FolioTF;
                        worksheet.Cells["BS" + index].Value = body.Items[c].NombresApellidos;
                    }

                    package.Save();
                    package.Stream.Position = 0;
                    var split = documentName.Split(".");
                    List<string> newSplit = new List<string>();
                    newSplit.Add(split[0]);
                    newSplit.Add(DateTime.Now.ToString("ddMMyyyyHHmmss"));
                    newSplit.Add(split[1]);
                    string newBlobName = newSplit[0] + "-" + newSplit[1] + "." + newSplit[2];
                    BlobContainerClient newContainerClient = blobServiceClient.GetBlobContainerClient(containerName + "history");
                    await newContainerClient.CreateIfNotExistsAsync(publicAccessType: PublicAccessType.BlobContainer);
                    BlobClient newBlobClient = newContainerClient.GetBlobClient(newBlobName);
                    await newBlobClient.StartCopyFromUriAsync(blob.Uri);
                    newBlobClient.Upload(package.Stream, overwrite: true);
                    var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    package.Stream.Position = 0;
                    return File(package.Stream, contentType, newBlobName);
                }

                return NotFound("No se encontró la hoja de Excel o no se leyeron datos.");


            }
            catch (Exception ex)
            {
                return BadRequest($"Error al leer el archivo Excel: {ex.Message}");
            }

        }

        [HttpGet("{containerName}/{documentName}")]
        public async Task<IActionResult> DownloadFile([FromHeader(Name = "ApiKey")] string apikey, string containerName, string documentName)
        {
            try
            {
                containerName = containerName.ToLower();

                BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                BlobClient blob = containerClient.GetBlobClient(documentName);
                BlobDownloadInfo blobDownloadInfo = await blob.DownloadAsync();
                MemoryStream memoryStream = new MemoryStream();
                await blobDownloadInfo.Content.CopyToAsync(memoryStream);
                byte[] blobBytes = memoryStream.ToArray();
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                return File(blobBytes, contentType, documentName);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al leer el archivo Excel: {ex.Message}");
            }

        }
    }
}
