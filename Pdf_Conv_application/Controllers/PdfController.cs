// Controllers/PdfController.cs
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Threading.Tasks;

[Route("api/[controller]")]
[ApiController]
public class PdfController : ControllerBase
{
    private readonly PdfConverterService _pdfConverterService;
    private readonly ILogger<PdfController> _logger;

    public PdfController(PdfConverterService pdfConverterService, ILogger<PdfController> logger)
    {
        _pdfConverterService = pdfConverterService;
        _logger = logger;
    }

    [HttpPost("html-file-to-pdf")]
    public async Task<IActionResult> ConvertHtmlFileToPdf(IFormFile htmlFile)
    {
        if (htmlFile == null || htmlFile.Length == 0)
        {
            _logger.LogError("Invalid HTML file.");
            return BadRequest("Invalid HTML file.");
        }

        using (var stream = new MemoryStream())
        {
            await htmlFile.CopyToAsync(stream);
            stream.Position = 0;
            try
            {
                var pdfBytes = _pdfConverterService.ConvertHtmlFileToPdf(stream);
                return File(pdfBytes, "application/pdf", "converted.pdf");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to convert HTML file to PDF.");
                return StatusCode(500, "An error occurred while converting HTML file to PDF.");
            }
        }
    }

    [HttpPost("excel-file-to-pdf")]
    public async Task<IActionResult> ConvertExcelFileToPdf(IFormFile excelFile)
    {
        if (excelFile == null || excelFile.Length == 0)
        {
            _logger.LogError("Invalid Excel file.");
            return BadRequest("Invalid Excel file.");
        }

        using (var stream = new MemoryStream())
        {
            await excelFile.CopyToAsync(stream);
            stream.Position = 0;
            try
            {
                var pdfBytes = _pdfConverterService.ConvertExcelToPdf(stream);
                return File(pdfBytes, "application/pdf", "converted.pdf");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to convert Excel file to PDF.");
                return StatusCode(500, "An error occurred while converting Excel file to PDF.");
            }
        }
    }

    [HttpPost("word-file-to-pdf")]
    public async Task<IActionResult> ConvertWordFileToPdf(IFormFile wordFile)
    {
        if (wordFile == null || wordFile.Length == 0)
        {
            _logger.LogError("Invalid Word file.");
            return BadRequest("Invalid Word file.");
        }

        using (var stream = new MemoryStream())
        {
            await wordFile.CopyToAsync(stream);
            stream.Position = 0;
            try
            {
                var pdfBytes = _pdfConverterService.ConvertWordToPdfWithImages(stream);
                return File(pdfBytes, "application/pdf", "converted.pdf");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to convert Word file to PDF with images.");
                return StatusCode(500, "An error occurred while converting Word file to PDF with images.");
            }
        }
    }
}
