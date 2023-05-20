namespace InteropNPOIBenchmark
{
    class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("\nInterop VS NPOI Benchmark");
            Console.Write("Creating Excel Files with 1000 rows and 10 columns: ... ");

            var path = @"D:\TemporaryFiles\";
            CreateExcelWorkbooks(20, path);
            
            Console.WriteLine("Done.");

            // NPOI
            for (int t = 0; t < 3; t++)
            {
                var npoiTimes = new List<long>();

                Console.WriteLine($"Test #{t+1}: ");
                
                for (int i = 0; i < 20; i++)
                {
                    npoiTimes.Add(Benchmark.NPOI(path + $"Excel{i+1}.xlsx"));
                }
                GetStatistics(npoiTimes);
            }

            // Interop
            for (int t = 0; t < 3; t++)
            {
                var interopTimes = new List<long>();

                Console.WriteLine($"Test #{t + 1}: ");

                for (int i = 0; i < 20; i++)
                {
                    interopTimes.Add(Benchmark.Interop(path + $"Excel{i + 1}.xlsx"));
                }
                GetStatistics(interopTimes);
            }

            Console.WriteLine("Done");
        }

        static void CreateExcelWorkbooks(int wbNum, string rootPath, int nbRows = 1000, int nbCols = 10)
        {
            for (int i = 0; i < wbNum; i++)
            {
                var wb = new XSSFWorkbook();
                var ws = wb.CreateSheet("Sheet1");

                for (int r = 0; r < nbRows; r++)
                {
                    var npoiRow = ws.CreateRow(r);
                    for (int c = 0; c < nbCols; c++)
                    {
                        var npoiCell = npoiRow.CreateCell(c);
                        npoiCell.SetCellValue("Cell " + r + " " + c);
                    }
                }
                using (var fs = new FileStream(rootPath + $"Excel{i+1}.xlsx", FileMode.Create))
                {
                    wb.Write(fs);
                }
            }
        }

        static void GetStatistics(List<long> benchmarkTimes)
        {
            Console.WriteLine("\tMin = " + Statistics.Min(benchmarkTimes) + "\tMax = " + Statistics.Max(benchmarkTimes) + "\tMean = " + Statistics.Mean(benchmarkTimes));
        }

    }
}
