using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using ExcelAndWordApi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.Internal;
using OfficeOpenXml;

namespace ExcelAndWordApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ContainersController : ControllerBase
    {
        string connectionString = "DefaultEndpointsProtocol=https;AccountName=riascosservicesstorage;AccountKey=ycHtf5e4s/n4dH4dCnr9ayVro8Ka0nywY9uqwJba50mn1LGfZM6CWI2pTclU3XWnSHom8oc5oc9L+AStqBHiKA==;EndpointSuffix=core.windows.net";

        [HttpGet]
        public async Task<IActionResult> getContainers([FromHeader(Name = "ApiKey")] string apikey)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
            var containers = blobServiceClient.GetBlobContainers();
            return Ok(containers);
        }

        [HttpGet("{containerName}")]
        public async Task<IActionResult> getContainer([FromHeader(Name = "ApiKey")] string apikey, string containerName)
        {
            BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
            var containers = blobServiceClient.GetBlobContainerClient(containerName);
            
            return Ok(new { Container = containers, Blobs = containers.GetBlobs().ToArray() });
        }
    }
}
