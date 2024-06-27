using ClosedXML.Excel;
using ExcelGenerator.Models;

namespace ExcelGenerator.Core;

internal class HeaderMethods
{
    internal static void TableHeaderCreation(IXLWorksheet worksheet, HashSet<Header> headers, string[] columns)
    {
        int index = 1;
        foreach (Header header in headers)
        {
            worksheet.Cell($"{columns[index - 1]}1").Value = header.Translation ?? header.ColumnName;
            index++;
        }

        worksheet.Range($"{columns.First()}:{columns.Last()}").Style.Font.FontSize = Constants.DefaultFontSize;
        worksheet.Range($"{columns.First()}:{columns.Last()}").Style.Font.Bold = true;
        worksheet.Range($"{columns.First()}:{columns.Last()}").Style.Font.FontName = Constants.DefaultFontName;
        worksheet.Range($"{columns.First()}:{columns.Last()}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        worksheet.Range($"{columns.First()}:{columns.Last()}").Style.Border.OutsideBorderColor = XLColor.Black;
    }
}
