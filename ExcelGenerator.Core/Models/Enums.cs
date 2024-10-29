namespace ExcelGenerator.Core;

public record Enums
{
    public enum FileFormat
    {
        Xlsx = 1,
        CsvComma = 2,
        CsvSemiColon = 3
    }

    public enum PageFormat
    {
        Table = 1,
        PaySlip = 2
    }

    public enum ImagePosition
    {
        TableRight = 1,
        TableBottom = 2
    }
}
