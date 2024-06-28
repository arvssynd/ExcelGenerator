using ExcelGenerator.Core;

namespace ExcelGenerator.Test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            #region DATA
            List<TestModel> model =
            [
                new TestModel
                {
                    HeaderString1 = "Hello",
                    HeaderBoolean4 = true,
                    HeaderDateTime5 = DateTime.Now,
                    HeaderDouble3   = 0.0,
                    HeaderInteger2 = 2,
                    HeaderCurrency6 = 1500
                },
                new TestModel
                {
                    HeaderString1 = "Prova",
                    HeaderBoolean4 = false,
                    HeaderDateTime5 = DateTime.Now,
                    HeaderDouble3 = 5.0,
                    HeaderInteger2 = 4,
                    HeaderCurrency6 = 1200.50m
                },
            ];
            #endregion

            List<Page<TestModel>> pages =
            [
                new Page<TestModel>
                {
                    Items = model,
                    CurrencyColumns = ["HeaderCurrency6"],
                    NumericColumns = ["HeaderDouble3"],
                    DateTimeColumns = ["HeaderDateTime5"],
                    DateTimeFormat = "dd/MM/yyyy HH:ss",
                    PageName = "Test1",
                    ExcludedColumns = ["HeaderBoolean4"]
                },
                new Page<TestModel>
                {
                    Items = model,
                    CurrencyColumns = ["HeaderCurrency6"],
                    NumericColumns = ["HeaderDouble3"],
                    DateTimeColumns = ["HeaderDateTime5"],
                    BooleanTrueTranslation = "Sì",
                    BooleanFalseTranslation = "No",
                    Headers =
                    [
                        new Header
                        {
                            ColumnName = "HeaderString1",
                            Translation = "TraduzioneHeaderString1"
                        },
                        new Header
                        {
                            ColumnName = "HeaderInteger2",
                            FontSize = 12
                        },
                        new Header
                        {
                            ColumnName = "HeaderDouble3"
                        },
                        new Header
                        {
                            ColumnName = "HeaderDateTime5"
                        },
                        new Header
                        {
                            ColumnName = "HeaderCurrency6"
                        },
                        new Header
                        {
                            ColumnName = "HeaderBoolean4"
                        },
                    ]
                }
            ];

            var excelStream = ExcelOrchestrator.GenerateExcel(pages);
            string localFilePath = $"{DateTime.Now:yyyyMMdd HHmmss}-Excel.xlsx";
            excelStream.Position = 0;
            using (FileStream fileStream = new(localFilePath, FileMode.Create, FileAccess.Write))
            {
                excelStream.CopyTo(fileStream);
            }

            Console.WriteLine("Excel file saved successfully.");
        }

        private class TestModel
        {
            public string HeaderString1 { get; set; }
            public int HeaderInteger2 { get; set; }
            public double HeaderDouble3 { get; set; }
            public bool HeaderBoolean4 { get; set; }
            public DateTime HeaderDateTime5 { get; set; }
            public decimal HeaderCurrency6 { get; set; }
            public long HeaderLongTest { get; set; }
        }
    }
}
