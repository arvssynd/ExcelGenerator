﻿using ClosedXML.Excel;

namespace ExcelGenerator.Core;

internal static class ValuesMethods
{
    internal static void TableValuesCreation<T>(this IXLWorksheet worksheet,
        Header[] headers,
        string[] columns,
        Page<T> items,
        HashSet<string> numericColumns,
        HashSet<string> currencyColumns,
        HashSet<string> dateTimeColumns,
        HashSet<string> excludedColumns,
        string? timeZone)
    {
        if (items is not null && items.Items.Count > 0)
        {
            for (int i = 1; i <= items.Items.Count; i++)
            {
                for (int j = 1; j <= headers.Length; j++)
                {
                    if (!excludedColumns.Any(x => x.Equals(headers[j - 1].ColumnName.ToUpper())) && !excludedColumns.Any(x => x.Equals(headers[j - 1].ColumnName)))
                    {
                        var property = items.Items[i - 1]?.GetType().GetProperty(headers[j - 1].ColumnName.ToUpper()) ?? items.Items[i - 1]?.GetType().GetProperty(headers[j - 1].ColumnName);
                        if (property is not null)
                        {
                            var type = property.PropertyType.Name;
                            if (type == Constants.DateTime || (property?.PropertyType?.GenericTypeArguments?.Any(x => x.Name == Constants.DateTime) ?? false))
                            {
                                var date = (DateTime?)property.GetValue(items.Items[i - 1]);
                                if (date.HasValue && !string.IsNullOrWhiteSpace(timeZone))
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
                            else if (Constants.Long.Contains(type) || (property?.PropertyType?.GenericTypeArguments?.Any(x => Constants.Long.Contains(type)) ?? false))
                            {
                                worksheet.Cell($"{columns[j - 1]}{i + 1}").Value = (long?)property?.GetValue(items.Items[i - 1]);
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
            }

            worksheet.Style.Font.FontSize = Constants.DefaultFontSize;
            worksheet.Style.Font.FontName = Constants.DefaultFontName;
            worksheet.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Style.Border.OutsideBorderColor = XLColor.Black;

            if (numericColumns.Count != 0)
            {
                foreach (var column in numericColumns)
                {
                    if (!excludedColumns.Any(x => x.Equals(column.ToUpper())) && !excludedColumns.Any(x => x.Equals(column)))
                    {
                        var _col = worksheet.ColumnsUsed(x => x.FirstCell().GetString() == column);
                        if (string.IsNullOrWhiteSpace(headers.FirstOrDefault(x => x.ColumnName == column)?.NumericFormat))
                        {
                            _col.Style.NumberFormat.Format = Constants.DefaultNumericFormat;
                        }
                        else
                        {
                            _col.Style.NumberFormat.Format = headers.FirstOrDefault(x => x.ColumnName == column)?.NumericFormat;
                        }
                    }
                }
            }

            if (currencyColumns.Count != 0)
            {
                foreach (var column in currencyColumns)
                {
                    if (!excludedColumns.Any(x => x.Equals(column.ToUpper())) && !excludedColumns.Any(x => x.Equals(column)))
                    {
                        var _col = worksheet.ColumnsUsed(x => x.FirstCell().GetString() == column);
                        if (string.IsNullOrWhiteSpace(headers.FirstOrDefault(x => x.ColumnName == column)?.CurrencyFormat))
                        {
                            _col.Style.NumberFormat.Format = Constants.DefaultCurrencyFormat;
                        }
                        else
                        {
                            _col.Style.NumberFormat.Format = headers.FirstOrDefault(x => x.ColumnName == column)?.CurrencyFormat;
                        }
                    }
                }
            }

            if (dateTimeColumns.Count != 0 && !string.IsNullOrWhiteSpace(items.DateTimeFormat))
            {
                foreach (var column in dateTimeColumns)
                {
                    if (!excludedColumns.Any(x => x.Equals(column.ToUpper())) && !excludedColumns.Any(x => x.Equals(column)))
                    {
                        var _col = worksheet.ColumnsUsed(x => x.FirstCell().GetString() == column);
                        _col.Style.DateFormat.Format = items.DateTimeFormat;
                    }
                }
            }

            if (headers.Any(x => !string.IsNullOrWhiteSpace(x.FontName) || x.FontSize.HasValue))
            {
                foreach (var column in headers.Where(x => !string.IsNullOrWhiteSpace(x.FontName) || x.FontSize.HasValue))
                {
                    if (!excludedColumns.Any(x => x.Equals(column.ColumnName.ToUpper())) && !excludedColumns.Any(x => x.Equals(column.ColumnName)))
                    {
                        var _col = worksheet.ColumnsUsed(x => x.FirstCell().GetString() == (column.Translation ?? column.ColumnName));
                        if (!string.IsNullOrWhiteSpace(column.FontName))
                        {
                            _col.Style.Font.FontName = column.FontName;
                        }

                        if (column.FontSize.HasValue)
                        {
                            _col.Style.Font.FontSize = column.FontSize ?? Constants.DefaultFontSize;
                        }
                    }
                }
            }
        }
    }
}
