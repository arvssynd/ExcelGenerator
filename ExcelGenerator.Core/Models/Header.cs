namespace ExcelGenerator.Core;

public class Header
{
    /// <summary>
    /// Rispecchia il nome della proprietà contenente il valore
    /// </summary>
    public required string ColumnName { get; set; }

    /// <summary>
    /// Nome dell'header tradotto, se non specificato viene inserito ColumnName
    /// </summary>
    public string? Translation { get; set; }

    /// <summary>
    /// Se popolato sostituisce le dimensioni del font per quella colonna, default 11
    /// </summary>
    public double? FontSize { get; set; }

    /// <summary>
    /// Se popolato sostituisce le dimensioni del font per quella colonna, default Aptos Narrow
    /// </summary>
    public string? FontName { get; set; }

    /// <summary>
    /// Se popolato sostituisce il formato numerico per quella colonna, default 0.0
    /// </summary>
    public string? NumericFormat { get; set; }

    /// <summary>
    /// Se popolato sostituisce il formato monetario per quella colonna, default #,##0.00 €
    /// </summary>
    public string? CurrencyFormat { get; set; }
}
