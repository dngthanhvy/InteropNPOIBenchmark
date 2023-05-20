using Microsoft.Office.Interop.Excel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Diagnostics;
using System.IO;

namespace InteropNPOIBenchmark;

public static class Benchmark
{
    public static long Interop(string filePath)
    {
        var stopwatch = Stopwatch.StartNew();
        var excelApp = new Application();

        excelApp.ScreenUpdating = false;
        excelApp.DisplayAlerts = false;

        try
        {
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = (Worksheet)workbook.Sheets["Sheet1"];
            var range = worksheet.UsedRange;
            var rowCount = range.Rows.Count;
            var colCount = range.Columns.Count;
                
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    var cell = range.Cells[row, col];
                }
            }

            workbook.Close(false);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
        finally
        {
            excelApp.Quit();
            stopwatch.Stop();
            Console.WriteLine("Execution time: " + stopwatch.Elapsed);
        }

        return stopwatch.ElapsedMilliseconds;
    }
    
    public static long NPOI(string filePath)
    {
        var stopwatch = Stopwatch.StartNew();
        try
        {
            // Open the Excel file
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Load the workbook
                var workbook = new XSSFWorkbook(fileStream);

                // Get the first worksheet
                var sheet = workbook.GetSheetAt(0);

                // Read data from the worksheet
                for (int row = 0; row <= sheet.LastRowNum; row++)
                {
                    var currentRow = sheet.GetRow(row);
                    if (currentRow == null)
                        continue;

                    for (int col = 0; col < currentRow.LastCellNum; col++)
                    {
                        var cell = currentRow.GetCell(col);
                    }
                }

                workbook.Close();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
        finally
        {
            stopwatch.Stop();
        }
        return stopwatch.ElapsedMilliseconds;
    }
}