using ClosedXML.Excel;

namespace ExcelGenerator.Core;

internal class HeaderMethods
{
    internal static void TableHeaderCreation(
        IXLWorksheet worksheet,
        HashSet<Header> headers,
        HashSet<string> excludedColumns,
        string[] columns)
    {
        int index = 1;
        foreach (Header header in headers)
        {
            if (!excludedColumns.Any(x => x.Equals(header.ColumnName.ToUpper())) && !excludedColumns.Any(x => x.Equals(header.ColumnName)))
            {
                worksheet.Cell($"{columns[index - 1]}").Value = header.Translation ?? header.ColumnName;
                index++;
            }
        }

        worksheet.Range($"{columns.First()}:{columns.Last()}").Style.Font.FontSize = Constants.DefaultFontSize;
        worksheet.Range($"{columns.First()}:{columns.Last()}").Style.Font.Bold = true;
        worksheet.Range($"{columns.First()}:{columns.Last()}").Style.Font.FontName = Constants.DefaultFontName;
        worksheet.Range($"{columns.First()}:{columns.Last()}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        worksheet.Range($"{columns.First()}:{columns.Last()}").Style.Border.OutsideBorderColor = XLColor.Black;
    }
}
