using ClosedXML.Excel;
using ExcelGenerator.Models;

namespace ExcelGenerator.Core;

public static class ExcelOrchestrator
{
    public static Stream GenerateExcel<T>(this List<Page<T>> data)
    {
        Stream fs = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            foreach (var page in data)
            {
                var worksheet = workbook.Worksheets.Add(page.PageName);
                string[] columns = Generate().Take(page.Headers.Count).ToArray();
                HeaderMethods.TableHeaderCreation(worksheet, page.Headers, columns);
                ValuesMethods.TableValuesCreation(
                    worksheet,
                    [.. page.Headers],
                    columns,
                    [.. data],
                    page.NumericColumns ?? [],
                    page.CurrencyColumns ?? [],
                    page.TimeZone);
            }

            workbook.SaveAs(fs);
        }

        fs.Position = 0;
        return fs;
    }

    private static string ToBase26(long i)
    {
        if (i == 0) return ""; i--;
        return ToBase26(i / 26) + (char)('A' + i % 26);
    }

    private static IEnumerable<string> Generate()
    {
        long n = 0;
        while (true) yield return ToBase26(++n);
    }
}
