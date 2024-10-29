using static ExcelGenerator.Core.Enums;

namespace ExcelGenerator.Core;

public class Image
{
    public string Base64 { get; set; }
    public ImagePosition Position { get; set; }
    public int? Width { get; set; }
    public int? Height { get; set; }
}
