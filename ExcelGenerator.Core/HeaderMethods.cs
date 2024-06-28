using ClosedXML.Excel;

namespace ExcelGenerator.Core;

internal static class HeaderMethods
{
    internal static void TableHeaderCreation(this IXLWorksheet worksheet,
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

    internal static void GenerateHeaders<T>(this Page<T> page)
    {
        if (page.Headers.Count == 0 && page.Items.FirstOrDefault()?.GetType().GetProperties() is not null)
        {
            page.Headers = page.Items.First()!.GetType().GetProperties().Select(x => new Header() { ColumnName = x.Name }).ToHashSet();
        }

        page.Headers = page.Headers.Where(x => !page.ExcludedColumns!.Contains(x.ColumnName)).ToHashSet();
    }
}
