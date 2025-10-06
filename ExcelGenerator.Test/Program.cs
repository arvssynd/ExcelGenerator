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
                    Items = [], // model, // your data set
                    BooleanTrueTranslation = "Sì",  // value displayed if the boolean is true, prints true if nothing is specified
                    BooleanFalseTranslation = "No", // value displayed if the boolean is false, prints false if nothing is specified
                    CurrencyColumns = ["HeaderCurrency6"],  // columns displayed in currency format
                    NumericColumns = ["HeaderDouble3"], // columns displayed in numeric format
                    DateTimeColumns = ["HeaderDateTime5"], // columns displayed in datetime format
                    DateTimeFormat = "dd/MM/yyyy HH:ss",    // date time format
                    PageName = "Test1", // excel page title
                    ExcludedColumns = ["HeaderBoolean4"],   // columns to exclude
                    Headers = // headers to print and specs, if headers is not specified all the Items properties will be displayed
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
                            ColumnName = "HeaderCurrency6",
                            CurrencyFormat = "example_format"
                        },
                        new Header
                        {
                            ColumnName = "HeaderBoolean4"
                        },
                    ],
                    TimeZone = "Europe/Rome" // timezone supported by TimeZoneInfo.FindSystemTimeZoneById
                },
                new Page<TestModel>
                {
                    // you can also create a page specifying only the dataset
                    Items = [], //model
                }
            ];

            // xlsx
            var excelStream = ExcelOrchestrator.GenerateExcel(pages);
            string localFilePath = $"{DateTime.Now:yyyyMMdd HHmmss}-Excel.xlsx";

            // csv
            //var excelStream = ExcelOrchestrator.GenerateExcel(pages, Enums.FileFormat.CsvComma);
            //string localFilePath = $"{DateTime.Now:yyyyMMdd HHmmss}-Excel.csv";

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
