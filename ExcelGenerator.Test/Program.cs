namespace ExcelGenerator.Test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<TestModel> model = new()
            {
                new TestModel 
                {

                }
            };
        }

        private class TestModel
        {
            public string HeaderString1 { get; set; }
            public int HeaderInteger2 { get; set; }
            public double HeaderDouble3 { get; set; }
            public bool HeaderBoolean4 { get; set; }
            public DateTime HeaderDateTime5 { get; set; }
        }
    }
}
