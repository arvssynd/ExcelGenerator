using ClosedXML.Excel;

namespace ExcelGenerator.Core;

public static class ExcelOrchestrator
{
    public static Stream GenerateExcel<T>(this List<Page<T>> data)
    {
        Stream fs = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var index = 1;
            foreach (var page in data)
            {
                var worksheet = workbook.Worksheets.Add(!string.IsNullOrWhiteSpace(page.PageName) ? page.PageName : $"Page {index}");
                if (page.Headers.Count == 0 && page.Items.FirstOrDefault()?.GetType().GetProperties() is not null)
                {
                    page.Headers = page.Items.First()!.GetType().GetProperties().Select(x => new Header() { ColumnName = x.Name }).ToHashSet();                    
                }
                page.Headers = page.Headers.Where(x => !page.ExcludedColumns!.Contains(x.ColumnName)).ToHashSet();
                string[] columns = [.. Generate(page.Headers.Count)];

                // header
                HeaderMethods.TableHeaderCreation(
                    worksheet,
                    page.Headers,
                    page.ExcludedColumns ?? [],
                    columns.Select(x => $"{x}1").ToArray());

                // values
                ValuesMethods.TableValuesCreation(
                    worksheet,
                    [.. page.Headers],
                    columns,
                    page,
                    page.NumericColumns ?? [],
                    page.CurrencyColumns ?? [],
                    page.DateTimeColumns ?? [],
                    page.ExcludedColumns ?? [],
                    page.TimeZone);

                index++;
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
