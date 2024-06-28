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
                // page
                var worksheet = workbook.Worksheets.Add(!string.IsNullOrWhiteSpace(page.PageName) ? page.PageName : $"Page {index}");

                switch (page.Format)
                {
                    case Enums.PageFormat.Table:
                        // header
                        page.GenerateHeaders();
                        string[] columns = [.. Generate(page.Headers.Count)];
                        worksheet.TableHeaderCreation(
                            page.Headers,
                            page.ExcludedColumns ?? [],
                            columns.Select(x => $"{x}1").ToArray()
                        );

                        // values
                        worksheet.TableValuesCreation(
                            [.. page.Headers],
                            columns,
                            page,
                            page.NumericColumns ?? [],
                            page.CurrencyColumns ?? [],
                            page.DateTimeColumns ?? [],
                            page.ExcludedColumns ?? [],
                            page.TimeZone
                        );

                        // TODO footer tabella
                        break;
                    case Enums.PageFormat.PaySlip:
                        break;
                    default:
                        break;
                }

                index++;
            }

            workbook.SaveAs(fs);
        }

        fs.Position = 0;
        return fs;
    }

    private static IEnumerable<string> Generate(int numberOfColumns)
    {
        for (int i = 0; i < numberOfColumns; i++)
        {
            yield return GetExcelColumnName(i + 1);
        }
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
