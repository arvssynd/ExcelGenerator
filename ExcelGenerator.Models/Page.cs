namespace ExcelGenerator.Models;

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
    /// Colonne numeriche
    /// </summary>
    public HashSet<string> NumericColumns { get; set; } = [];

    /// <summary>
    /// Colonne in formato monetario
    /// </summary>
    public HashSet<string> CurrencyColumns { get; set; } = [];

    /// <summary>
    /// Timezone in cui convertire le date
    /// </summary>
    public string? TimeZone { get; set; }

    /// <summary>
    /// Formato dell'excel
    /// 1. Tabella tradizionale
    /// 2. Busta paga
    /// ...
    /// </summary>
    public Enums.PageFormat Format { get; set; } = Enums.PageFormat.Table;
}
