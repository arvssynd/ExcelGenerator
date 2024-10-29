using ClosedXML.Excel;

namespace ExcelGenerator.Core;

internal static class ImageMethods
{
    internal static void AddImage(this IXLWorksheet worksheet, string base64Map, string cellCoordinates, int? width = null, int? height = null)
    {
        if (!string.IsNullOrWhiteSpace(base64Map))
        {
            var bytes = Convert.FromBase64String(base64Map);
            var contents = new MemoryStream(bytes);

            if (!width.HasValue && !height.HasValue)
            {
                worksheet
                    .AddPicture(contents)
                    .MoveTo(worksheet.Cell(cellCoordinates));
            }
            else
            {
                worksheet
                    .AddPicture(contents)
                    .MoveTo(worksheet.Cell(cellCoordinates))
                    .WithSize(width.Value, height.Value);
            }
        }
    }
}
