using ClosedXML.Excel;
using System.Text;

namespace ExcelGenerator.Core;

public static class ExcelOrchestrator
{
    public static Stream GenerateExcel<T>(this List<Page<T>> data, Enums.FileFormat fileFormat = Enums.FileFormat.Xlsx)
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

                        // image
                        foreach (var img in page.Images)
                        {
                            switch (img.Position)
                            {
                                case Enums.ImagePosition.TableRight:
                                    worksheet.AddImage(img.Base64, $"{columns.Last()}2");
                                    break;
                                case Enums.ImagePosition.TableBottom:
                                    break;
                                default:
                                    break;
                            }
                        }

                        // TODO footer tabella
                        break;
                    case Enums.PageFormat.PaySlip:
                        break;
                    default:
                        break;
                }

                index++;
            }

            //workbook.SaveAs(fs);

            switch (fileFormat)
            {
                case Enums.FileFormat.CsvComma:
                    workbook.TransformXlsxStreamToCsvStream(fs, ",");
                    break;
                case Enums.FileFormat.CsvSemiColon:
                    workbook.TransformXlsxStreamToCsvStream(fs);
                    break;
                case Enums.FileFormat.Xlsx:
                default:
                    workbook.SaveAs(fs);
                    fs.Position = 0;
                    break;
            }
        }

        return fs;
    }

    private static Stream TransformXlsxStreamToCsvStream(this XLWorkbook workbook, Stream memoryStream, string separator = ";")
    {
        using (var writer = new StreamWriter(memoryStream, encoding: Encoding.UTF8, leaveOpen: true))
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    var line = string.Join(separator, row.Cells(1, row.LastCellUsed().Address.ColumnNumber).Select(cell => cell.GetValue<string>()));
                    writer.WriteLine(line);
                }

                writer.WriteLine();
            }
        }

        return memoryStream;
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
