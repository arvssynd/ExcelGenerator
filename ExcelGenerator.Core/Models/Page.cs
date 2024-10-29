namespace ExcelGenerator.Core;

public class Page<T>
{
    /// <summary>
    /// Dati
    /// </summary>
    public List<T> Items { get; set; } = [];

    /// <summary>
    /// Nome della pagina
    /// </summary>
    public string? PageName { get; set; }

    /// <summary>
    /// Configurazione degli header
    /// </summary>
    public HashSet<Header> Headers { get; set; } = [];

    /// <summary>
    /// Colonne numeriche non intere
    /// </summary>
    public HashSet<string> NumericColumns { get; set; } = [];

    /// <summary>
    /// Colonne in formato monetario
    /// </summary>
    public HashSet<string> CurrencyColumns { get; set; } = [];

    /// <summary>
    /// Colonne in formato data
    /// </summary>
    public HashSet<string> DateTimeColumns { get; set; } = [];

    /// <summary>
    /// Se non si specificano le colonne si può specificare quali escludere
    /// </summary>
    public HashSet<string> ExcludedColumns { get; set; } = [];

    /// <summary>
    /// Timezone in cui convertire le date
    /// </summary>
    public string? TimeZone { get; set; }

    /// <summary>
    /// Formato delle date
    /// </summary>
    public string? DateTimeFormat { get; set; }

    /// <summary>
    /// Valore da mostrare se il campo booleano è true
    /// </summary>
    public string? BooleanTrueTranslation { get; set; } = "True";

    /// <summary>
    /// Valore da mostrare se il campo booleano è false
    /// </summary>
    public string? BooleanFalseTranslation { get; set; } = "False";

    /// <summary>
    /// Formato dell'excel
    /// 1. Tabella tradizionale
    /// 2. Busta paga
    /// ...
    /// </summary>
    public Enums.PageFormat Format { get; set; } = Enums.PageFormat.Table;

    public HashSet<Image> Images { get; set; } = [];
}
