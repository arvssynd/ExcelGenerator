namespace ExcelGenerator.Core;

public record Constants
{
    public const string DefaultFontName = "Aptos Narrow";
    public const int DefaultFontSize = 11;
    public const string DefaultNumericFormat = "0.0";
    public const string DefaultCurrencyFormat = "#,##0.00 €";

    public const string DateTime = "DateTime";
    public static string[] Integer = ["Integer", "Int32"];
    public const string Decimal = "Decimal";
    public const string Double = "Double";
    public const string Boolean = "Boolean";
    public static string[] Long = ["Long", "Int64"];
}
