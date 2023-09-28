using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using ExcelAndWordApi.Models;
using ExcelAndWordApi.Models.Word;
using Microsoft.AspNetCore.Mvc;
using Xceed.Words.NET;

namespace ExcelAndWordApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WordController : ControllerBase
    {
        string connectionString = "DefaultEndpointsProtocol=https;AccountName=riascosservicesstorage;AccountKey=ycHtf5e4s/n4dH4dCnr9ayVro8Ka0nywY9uqwJba50mn1LGfZM6CWI2pTclU3XWnSHom8oc5oc9L+AStqBHiKA==;EndpointSuffix=core.windows.net";



        [HttpPost("{containerName}/{documentName}")]
        public async Task<IActionResult> Index([FromHeader(Name = "ApiKey")] string apikey, string containerName, string documentName, [FromBody] Doc001 body)
        {
            try
            {
                containerName = containerName.ToLower();
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                BlobClient blob = containerClient.GetBlobClient(documentName);
                BlobDownloadInfo blobDownloadInfo = await blob.DownloadAsync();
                MemoryStream memoryStream = new MemoryStream();
                await blobDownloadInfo.Content.CopyToAsync(memoryStream);
                var data = new List<List<object>>();
                memoryStream.Position = 0;
                using (var doc = DocX.Load(memoryStream))
                {
                    string[] textToBeReplaced = new string[] {
                "[numero_expediente]",
                "[fecha_expediente]",
                "[nit]",
                "[numero_nit]",
                "[empresa]",
                "[email_empresa]",
                "[dirección_empresa]",
                "[ciudad]",
                "[departamento]",
                "[pais]",
                "[numero_oficio]",
                "[numero_folios]",
                "[articulo]",
                "[ley]",
                "[campo_registro_fotografico]",
                "[nit_empresa]"
                };
                    string[] textToReplace = new string[] {
                    body.Expediente,
                    body.FechaExpediente.ToShortDateString(),
                    body.Nit,
                    body.NoNit,
                    body.Empresa,
                    body.EmpresaEmail,
                    body.EmpresaDireccion,
                    body.Ciudad,
                    body.Departamento,
                    body.Pais,
                    body.NoOficio,
                    body.NoFolio,
                    body.Articulo,
                    body.Ley,
                    body.CampoRegistroFotografico,
                    body.NitEmpresa
                };
                    for (var c = 0; c < textToBeReplaced.Length; c++)
                    {
                        doc.ReplaceText(textToBeReplaced[c], textToReplace[c]);
                    }
                    try
                    {
                        var stream = new MemoryStream();
                        doc.SaveAs(stream);

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
                        stream.Position = 0;
                        newBlobClient.Upload(stream, overwrite: true);
                        var contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                        stream.Position = 0;
                        return File(stream, contentType, newBlobName);
                    } catch(Exception ex)
                    {
                        return BadRequest($"Error al leer el archivo Word: {ex.Message}");
                    }
                    

                }

                return NotFound("No se encontró el archivo word.");


            }
            catch (Exception ex)
            {
                return BadRequest($"Error al leer el archivo Word: {ex.Message}");
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
                var contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                return File(blobBytes, contentType, documentName);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al leer el archivo Excel: {ex.Message}");
            }

        }


    }
}
