// Services/PdfConverterService.cs
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using ColorMode = DinkToPdf.ColorMode;

public class PdfConverterService
{
    private readonly IConverter _converter;
    private readonly ILogger<PdfConverterService> _logger;

    public PdfConverterService(ILogger<PdfConverterService> logger)
    {
        _converter = new SynchronizedConverter(new PdfTools());
        _logger = logger;
    }

    public byte[] ConvertHtmlFileToPdf(Stream htmlFile)
    {
        try
        {
            using (StreamReader reader = new StreamReader(htmlFile))
            {
                string htmlContent = reader.ReadToEnd();
                return ConvertHtmlToPdf(htmlContent);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error converting HTML file to PDF");
            throw;
        }
    }

    public byte[] ConvertHtmlToPdf(string htmlContent)
    {
        try
        {
            var globalSettings = new GlobalSettings
            {
                ColorMode = ColorMode.Color,
                Orientation = Orientation.Portrait,
                PaperSize = PaperKind.A4
            };

            var objectSettings = new ObjectSettings
            {
                HtmlContent = htmlContent,
                WebSettings = { DefaultEncoding = "utf-8" }
            };

            var pdf = new HtmlToPdfDocument
            {
                GlobalSettings = globalSettings,
                Objects = { objectSettings }
            };

            return _converter.Convert(pdf);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error converting HTML content to PDF");
            throw;
        }
    }

    public byte[] ConvertExcelToPdf(Stream excelFile)
    {
        try
        {
            using (var stream = new MemoryStream())
            {
                excelFile.CopyTo(stream);
                stream.Position = 0;

                IWorkbook workbook = new XSSFWorkbook(stream);
                ISheet sheet = workbook.GetSheetAt(0);

                StringBuilder htmlBuilder = new StringBuilder();
                htmlBuilder.Append("<html><body><table border='1'>");

                for (int i = 0; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    htmlBuilder.Append("<tr>");
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        htmlBuilder.Append(string.Format("<td>{0}</td>", cell.ToString()));
                    }
                    htmlBuilder.Append("</tr>");
                }

                htmlBuilder.Append("</table></body></html>");
                return ConvertHtmlToPdf(htmlBuilder.ToString());
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error converting Excel file to PDF");
            throw;
        }
    }

    public byte[] ConvertWordToPdfWithImages(Stream wordFile)
    {
        try
        {
            using (var doc = WordprocessingDocument.Open(wordFile, false))
            {
                var images = ConvertWordToImages(doc);
                return CreatePdfFromImages(images);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error converting Word file to PDF with images.");
            throw;
        }
    }

    private List<Image> ConvertWordToImages(WordprocessingDocument doc)
    {
        List<Image> images = new List<Image>();

        // Extract text from the Word document and create images.
        foreach (var element in doc.MainDocumentPart.Document.Body.Elements())
        {
            var image = new Bitmap(600, 800);
            using (Graphics g = Graphics.FromImage(image))
            {
                g.Clear(Color.White);
                g.DrawString(element.InnerText, new Font("Arial", 12), Brushes.Black, new RectangleF(10, 10, 580, 780));
            }
            images.Add(image);
        }
        return images;
    }

    private byte[] CreatePdfFromImages(List<Image> images)
    {
        var globalSettings = new GlobalSettings
        {
            ColorMode = ColorMode.Color,
            Orientation = Orientation.Portrait,
            PaperSize = PaperKind.A4
        };

        StringBuilder htmlBuilder = new StringBuilder();
        htmlBuilder.Append("<html><body>");

        foreach (var image in images)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, ImageFormat.Png);
                string base64Image = Convert.ToBase64String(ms.ToArray());
                htmlBuilder.Append($"<img src='data:image/png;base64,{base64Image}' style='width:100%;' />");
            }
        }

        htmlBuilder.Append("</body></html>");

        var objectSettings = new ObjectSettings
        {
            HtmlContent = htmlBuilder.ToString(),
            WebSettings = { DefaultEncoding = "utf-8" }
        };

        var pdf = new HtmlToPdfDocument
        {
            GlobalSettings = globalSettings,
            Objects = { objectSettings }
        };

        return _converter.Convert(pdf);
    }
}
