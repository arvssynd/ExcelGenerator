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
                string[] columns = [.. Generate(page.Headers.Count)];
                HeaderMethods.TableHeaderCreation(worksheet, page.Headers, columns.Select(x => $"{x}1").ToArray());
                ValuesMethods.TableValuesCreation(
                    worksheet,
                    [.. page.Headers],
                    columns,
                    page,
                    page.NumericColumns ?? [],
                    page.CurrencyColumns ?? [],
                    page.DateTimeColumns ?? [],
                    page.TimeZone);
            }

            workbook.SaveAs(fs);
        }

        fs.Position = 0;
        return fs;
    }

    private static List<string> Generate(int numberOfColumns)
    {
        var columns = new List<string>();
        for (int i = 0; i < numberOfColumns; i++)
        {
            columns.Add(GetExcelColumnName(i + 1));
        }

        return columns;
    }

    private static string GetExcelColumnName(int columnIndex)
    {
        string columnName = string.Empty;
        while (columnIndex > 0)
        {
            int modulo = (columnIndex - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            columnIndex = (columnIndex - 1) / 26;
        }
        return columnName;
    }
}
