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
        public IActionResult getWordFile([FromBody] Document body)
        {
            // Ruta al archivo Word existente
            var rutaArchivo = Path.Combine(filePath, fileName);

            // Abrir el documento Word
            using (var doc = DocX.Load(rutaArchivo))
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
                for(var c = 0; c < textToBeReplaced.Length; c++)
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
    }
}
