using ClosedXML.Excel;
using ExcelGenerator.Models;

namespace ExcelGenerator.Core;

internal class ValuesMethods
{
    internal static void TableValuesCreation<T>(
        IXLWorksheet worksheet,
        Header[] headers,
        string[] columns,
        T[] items,
        HashSet<string> numericColumns,
        HashSet<string> currencyColumns,
        string? timeZone)
    {
        if (items is not null && items.Length > 0)
        {
            for (int i = 1; i <= items.Length; i++)
            {
                for (int j = 1; j <= headers.Length; j++)
                {
                    var property = items[i - 1]?.GetType().GetProperty(headers[j - 1].ColumnName.ToUpper()) ?? items[i - 1]?.GetType().GetProperty(headers[j - 1].ColumnName);
                    if (property is not null)
                    {
                        var type = property.PropertyType.Name;
                        if (type == Constants.DateTime || (property?.PropertyType?.GenericTypeArguments?.Any(x => x.Name == Constants.DateTime) ?? false))
                        {
                            var date = (DateTime?)property.GetValue(items[i - 1]);
                            if (!string.IsNullOrWhiteSpace(timeZone) && date.HasValue)
                            {
                                date = TimeZoneInfo.ConvertTime(date.Value, TimeZoneInfo.FindSystemTimeZoneById(timeZone));
                            }

                            worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = date;
                        }
                        else if (type == Constants.Decimal || (property?.PropertyType?.GenericTypeArguments?.Any(x => x.Name == Constants.Decimal) ?? false))
                        {
                            worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = (decimal?)property?.GetValue(items[i - 1]);
                        }
                        else if (type == Constants.Integer || (property?.PropertyType?.GenericTypeArguments?.Any(x => x.Name == Constants.Integer) ?? false))
                        {
                            worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = (int?)property?.GetValue(items[i - 1]);
                        }
                        else
                        {
                            worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = property?.GetValue(items[i - 1])?.ToString();
                        }

                        if (numericColumns.Any(x => x == property!.Name))
                        {
                            if (string.IsNullOrWhiteSpace(headers[j - 1].NumericFormat))
                            {
                                worksheet.Cell($"{columns[j - 1]}{i + 1}").Style.NumberFormat.Format = "0.0";
                            }
                            else
                            {
                                worksheet.Cell($"{columns[j - 1]}{i + 1}").Style.NumberFormat.Format = headers[j - 1].NumericFormat;
                            }
                        }

                        if (currencyColumns.Any(x => x == property!.Name))
                        {
                            if (string.IsNullOrWhiteSpace(headers[j - 1].CurrencyFormat))
                            {
                                worksheet.Cell($"{columns[j - 1]}{i + 1}").Style.NumberFormat.Format = "#,##0.00 €";
                            }
                            else
                            {
                                worksheet.Cell($"{columns[j - 1]}{i + 1}").Style.NumberFormat.Format = headers[j - 1].CurrencyFormat;
                            }
                        }
                    }
                }
            }

            worksheet.Style.Font.FontSize = Constants.DefaultFontSize;
            worksheet.Style.Font.FontName = Constants.DefaultFontName;
            worksheet.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Style.Border.OutsideBorderColor = XLColor.Black;

            if (headers.Any(x => !string.IsNullOrWhiteSpace(x.FontName) || x.FontSize.HasValue))
            {
                for (int i = 1; i <= headers.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(headers[i - 1].FontName))
                    {
                        worksheet.Cell($"{columns[i - 1]}1").Style.Font.FontName = headers[i - 1].FontName;
                    }

                    if (headers[i - 1].FontSize.HasValue)
                    {
                        worksheet.Cell($"{columns[i - 1]}1").Style.Font.FontSize = headers[i - 1].FontSize ?? 0;
                    }
                }
            }
        }
    }
}
