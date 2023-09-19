using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using ExcelAndWordApi.Models;
using Microsoft.AspNetCore.Mvc;
using Xceed.Words.NET;

namespace ExcelAndWordApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WordController : ControllerBase
    {
        string filePath = Path.Combine("C:", "Word");
        string fileName = "Document.docx";
        [HttpPost]
        public async Task<IActionResult> getWordFile(string documentId, [FromBody] Document body)
        {
            // Ruta al archivo Word existente
            var rutaArchivo = Path.Combine(filePath, fileName);
            try
            {
                string connectionString = "DefaultEndpointsProtocol=https;AccountName=riascosservicesstorage;AccountKey=ycHtf5e4s/n4dH4dCnr9ayVro8Ka0nywY9uqwJba50mn1LGfZM6CWI2pTclU3XWnSHom8oc5oc9L+AStqBHiKA==;EndpointSuffix=core.windows.net"; // Debes reemplazar con tu cadena de conexión real
                string containerName = "generalcontainer";
                string blobName = documentId;

                BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                BlobClient blobClient = containerClient.GetBlobClient(blobName);
                var exits = await blobClient.ExistsAsync();
                if (exits)
                {
                    BlobDownloadInfo blobDownloadInfo = await blobClient.DownloadAsync();
                    MemoryStream memoryStream = new MemoryStream();
                    await blobDownloadInfo.Content.CopyToAsync(memoryStream);
                    // Abrir el documento Word
                    using (var doc = DocX.Load(memoryStream))
                    {
                        string[] textToBeReplaced = new string[] {
                "[customer]",
                "[supplier]",
                "[amount]",
                "[date]"
                };
                        string[] textToReplace = new string[] {
                    body.Customer,
                    body.Supplier,
                    body.Amount.ToString("c"),
                    DateTime.Now.ToShortDateString(),
                };
                        for (var c = 0; c < textToBeReplaced.Length; c++)
                        {
                            doc.ReplaceText(textToBeReplaced[c], textToReplace[c]);
                        }

                        // Guardar el documento modificado en memoria
                        using (var stream = new MemoryStream())
                        {
                            doc.SaveAs(stream);

                            // Convertir el documento a bytes
                            var bytes = stream.ToArray();

                            // Devolver el archivo Word como respuesta HTTP
                            return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Factura.docx");
                        }
                    }
                }
                else
                {
                    return NotFound($"No se encontro el archivo");
                }
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al leer el archivo Excel: {ex.Message}");
            }




                }


            }
        }
