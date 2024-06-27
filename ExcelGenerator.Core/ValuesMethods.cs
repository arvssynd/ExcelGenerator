using ClosedXML.Excel;
using ExcelGenerator.Models;

namespace ExcelGenerator.Core;

internal class ValuesMethods
{
    internal static void TableValuesCreation<T>(
        IXLWorksheet worksheet,
        Header[] headers,
        string[] columns,
        Page<T> items,
        HashSet<string> numericColumns,
        HashSet<string> currencyColumns,
        HashSet<string> dateTimeColumns,
        string? timeZone)
    {
        if (items is not null && items.Items.Count > 0)
        {
            for (int i = 1; i <= items.Items.Count; i++)
            {
                for (int j = 1; j <= headers.Length; j++)
                {
                    var property = items.Items[i - 1]?.GetType().GetProperty(headers[j - 1].ColumnName.ToUpper()) ?? items.Items[i - 1]?.GetType().GetProperty(headers[j - 1].ColumnName);
                    if (property is not null)
                    {
                        var type = property.PropertyType.Name;
                        if (type == Constants.DateTime || (property?.PropertyType?.GenericTypeArguments?.Any(x => x.Name == Constants.DateTime) ?? false))
                        {
                            var date = (DateTime?)property.GetValue(items.Items[i - 1]);
                            if (!string.IsNullOrWhiteSpace(timeZone) && date.HasValue)
                            {
                                date = TimeZoneInfo.ConvertTime(date.Value, TimeZoneInfo.FindSystemTimeZoneById(timeZone));
                            }

                            worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = date;
                        }
                        else if (type == Constants.Decimal || (property?.PropertyType?.GenericTypeArguments?.Any(x => x.Name == Constants.Decimal) ?? false))
                        {
                            worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = (decimal?)property?.GetValue(items.Items[i - 1]);
                        }
                        else if (type == Constants.Double || (property?.PropertyType?.GenericTypeArguments?.Any(x => x.Name == Constants.Double) ?? false))
                        {
                            worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = (double?)property?.GetValue(items.Items[i - 1]);
                        }
                        else if (Constants.Integer.Contains(type) || (property?.PropertyType?.GenericTypeArguments?.Any(x => Constants.Integer.Contains(type)) ?? false))
                        {
                            worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = (int?)property?.GetValue(items.Items[i - 1]);
                        }
                        else if (type == Constants.Boolean || (property?.PropertyType?.GenericTypeArguments?.Any(x => x.Name == Constants.Boolean) ?? false))
                        {
                            worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = (bool?)property?.GetValue(items.Items[i - 1]) ?? false ? items.BooleanTrueTranslation : items.BooleanFalseTranslation;
                        }
                        else
                        {
                            worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = property?.GetValue(items.Items[i - 1])?.ToString();
                        }
                    }
                }
            }

            worksheet.Style.Font.FontSize = Constants.DefaultFontSize;
            worksheet.Style.Font.FontName = Constants.DefaultFontName;
            worksheet.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Style.Border.OutsideBorderColor = XLColor.Black;

            if (numericColumns.Count != 0)
            {
                foreach (var column in numericColumns)
                {
                    var _col = worksheet.ColumnsUsed(x => x.FirstCell().GetString() == column);
                    if (string.IsNullOrWhiteSpace(headers.FirstOrDefault(x => x.ColumnName == column)?.NumericFormat))
                    {
                        _col.Style.NumberFormat.Format = "0.0";
                    }
                    else
                    {
                        _col.Style.NumberFormat.Format = headers.FirstOrDefault(x => x.ColumnName == column)?.NumericFormat;
                    }
                }
            }

            if (currencyColumns.Count != 0)
            {
                foreach (var column in currencyColumns)
                {
                    var _col = worksheet.ColumnsUsed(x => x.FirstCell().GetString() == column);
                    if (string.IsNullOrWhiteSpace(headers.FirstOrDefault(x => x.ColumnName == column)?.CurrencyFormat))
                    {
                        _col.Style.NumberFormat.Format = "#,##0.00 €";
                    }
                    else
                    {
                        _col.Style.NumberFormat.Format = headers.FirstOrDefault(x => x.ColumnName == column)?.CurrencyFormat;
                    }
                }
            }

            if (dateTimeColumns.Count != 0 && !string.IsNullOrWhiteSpace(items.DateTimeFormat))
            {
                foreach (var column in dateTimeColumns)
                {
                    var _col = worksheet.ColumnsUsed(x => x.FirstCell().GetString() == column);
                    _col.Style.DateFormat.Format = items.DateTimeFormat;
                }
            }

            if (headers.Any(x => !string.IsNullOrWhiteSpace(x.FontName) || x.FontSize.HasValue))
            {
                foreach (var column in headers.Where(x => !string.IsNullOrWhiteSpace(x.FontName) || x.FontSize.HasValue))
                {
                    var _col = worksheet.ColumnsUsed(x => x.FirstCell().GetString() == (column.Translation ?? column.ColumnName));
                    if (!string.IsNullOrWhiteSpace(column.FontName))
                    {
                        _col.Style.Font.FontName = column.FontName;
                    }

                    if (column.FontSize.HasValue)
                    {
                        _col.Style.Font.FontSize = column.FontSize ?? 0;
                    }
                }
            }
        }
    }
}
