using DinkToPdf;
using DinkToPdf.Contracts;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text;

public class PdfConverterService
{
    private readonly IConverter _converter;

    public PdfConverterService()
    {
        _converter = new SynchronizedConverter(new PdfTools());
    }

    public byte[] ConvertHtmlFileToPdf(Stream htmlFile)
    {
        using (StreamReader reader = new StreamReader(htmlFile))
        {
            string htmlContent = reader.ReadToEnd();
            return ConvertHtmlToPdf(htmlContent);
        }
    }

    public byte[] ConvertHtmlToPdf(string htmlContent)
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

    public byte[] ConvertExcelToPdf(Stream excelFile)
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
}
