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
        string sheetName = "Hoja1";
        string connectionString = "DefaultEndpointsProtocol=https;AccountName=riascosservicesstorage;AccountKey=ycHtf5e4s/n4dH4dCnr9ayVro8Ka0nywY9uqwJba50mn1LGfZM6CWI2pTclU3XWnSHom8oc5oc9L+AStqBHiKA==;EndpointSuffix=core.windows.net";
        string containerName = "generalcontainer";

        //[HttpPost]
        //public async Task<IActionResult> Index(string documentId, [FromBody] Person body)
        //{
        //    try
        //    {
        //        var blob = await getBlobClient(connectionString, containerName, documentId);
        //        Stream memoryStream = blob.DownloadStreamingAsync().Result.Value.Content;
        //        var package = new ExcelPackage(memoryStream);
        //        var worksheet = package.Workbook.Worksheets[sheetName]; // Cambia "MiHoja" al nombre de tu hoja
        //        var indextoInsert = worksheet.Dimension.End.Row + 1;
        //        worksheet.InsertRow(indextoInsert, 1);
        //        worksheet.Cells["A" + indextoInsert].Value = body.Name;
        //        worksheet.Cells["B" + indextoInsert].Value = body.Age;
        //        package.Save();
        //        uploadBlob(blob, package.Stream);

        //        return Ok();
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }

        //}
        //[HttpPut("{index}")]
        //public async Task<IActionResult> Index(string documentId, int index, [FromBody] Person body)
        //{
        //    var blob = await getBlobClient(connectionString, containerName, documentId);
        //    Stream memoryStream = blob.DownloadStreamingAsync().Result.Value.Content;
        //    var package = new ExcelPackage(memoryStream);
        //    var worksheet = package.Workbook.Worksheets[sheetName]; // Cambia "MiHoja" al nombre de tu hoja
        //    var rowToUpdate = worksheet.Cells["A" + index + ":B" + index];
        //    rowToUpdate["A" + index].Value = body.Name;
        //    rowToUpdate["B" + index].Value = body.Age;
        //    package.Save();
        //    uploadBlob(blob, package.Stream);
        //    return Ok();

        //}

        //[HttpDelete("{index}")]
        //public async Task<IActionResult> Index(string documentId, int index)
        //{
        //    var blob = await getBlobClient(connectionString, containerName, documentId);
        //    Stream memoryStream = blob.DownloadStreamingAsync().Result.Value.Content;
        //    var package = new ExcelPackage(memoryStream);
        //    var worksheet = package.Workbook.Worksheets[sheetName]; // Cambia "MiHoja" al nombre de tu hoja
        //    worksheet.DeleteRow(index);
        //    package.Save();
        //    uploadBlob(blob, package.Stream);
        //    return Ok();
        //}



        [HttpPost("{containerName}/{documentName}")]
        public async Task<IActionResult> Index(string containerName, string documentName, [FromBody] FT_FD_2558 body)
        {
            try
            {
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
                    newSplit.Add(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                    newSplit.Add(split[1]);
                    string newBlobName = newSplit[0] + newSplit[1] + newSplit[2];
                    BlobContainerClient newContainerClient = blobServiceClient.GetBlobContainerClient(containerName + "_history");
                    await newContainerClient.CreateIfNotExistsAsync();
                    BlobClient newBlobClient = newContainerClient.GetBlobClient("");
                    await newBlobClient.StartCopyFromUriAsync(blob.Uri);
                    newBlobClient.Upload(package.Stream, overwrite: true);
                    MemoryStream _memoryStream = new MemoryStream();
                    BlobDownloadInfo blobDownloadInfo = await newBlobClient.DownloadAsync();
                    await blobDownloadInfo.Content.CopyToAsync(memoryStream);
                    byte[] blobBytes = _memoryStream.ToArray();
                    var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    return File(blobBytes, contentType, newBlobName);
                }

                return NotFound("No se encontró la hoja de Excel o no se leyeron datos.");


            }
            catch (Exception ex)
            {
                return BadRequest($"Error al leer el archivo Excel: {ex.Message}");
            }

        }

        [HttpGet("{containerName}/{documentName}")]
        public async Task<IActionResult> DownloadFile(string containerName, string documentName)
        {
            try
            {
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
