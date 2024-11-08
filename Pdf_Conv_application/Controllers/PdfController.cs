using Microsoft.AspNetCore.Mvc;

[Route("api/[controller]")]
[ApiController]
public class PdfController : ControllerBase
{
    private readonly PdfConverterService _pdfConverterService;

    public PdfController(PdfConverterService pdfConverterService)
    {
        _pdfConverterService = pdfConverterService;
    }

    [HttpPost("html-file-to-pdf")]
    public async Task<IActionResult> ConvertHtmlFileToPdf(IFormFile htmlFile)
    {
        if (htmlFile == null || htmlFile.Length == 0)
            return BadRequest("Invalid HTML file.");

        using (var stream = new MemoryStream())
        {
            await htmlFile.CopyToAsync(stream);
            stream.Position = 0;
            var pdfBytes = _pdfConverterService.ConvertHtmlFileToPdf(stream);
            return File(pdfBytes, "application/pdf", "converted.pdf");
        }
    }

    [HttpPost("excel-file-to-pdf")]
    public async Task<IActionResult> ConvertExcelFileToPdf(IFormFile excelFile)
    {
        if (excelFile == null || excelFile.Length == 0)
            return BadRequest("Invalid Excel file.");

        using (var stream = new MemoryStream())
        {
            await excelFile.CopyToAsync(stream);
            stream.Position = 0;
            var pdfBytes = _pdfConverterService.ConvertExcelToPdf(stream);
            return File(pdfBytes, "application/pdf", "converted.pdf");
        }
    }
}
