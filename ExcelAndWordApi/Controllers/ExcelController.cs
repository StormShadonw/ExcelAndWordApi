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

        [HttpPost]
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
        [HttpGet]
        public IActionResult Index()
        {
            try
            {
                List<string> datos = new List<string>();

                string filePath = Path.Combine(Directory.GetCurrentDirectory(), "ArchivosExcel.xlsx");
                Application excelApp = new Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1]; // Hoja 1

                // Suponiendo que quieres leer datos en un rango específico, por ejemplo, A1:B10
                Excel.Range range = worksheet.Range["A1:B10"];

                foreach (Excel.Range cell in range)
                {
                    // Agregar el valor de cada celda a la lista de datos
                    datos.Add(cell.Value.ToString());
                }

                // Cerrar el archivo Excel sin guardar cambios
                workbook.Close(false);
                excelApp.Quit();
                return Ok(datos);
            } catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }
    }
}
