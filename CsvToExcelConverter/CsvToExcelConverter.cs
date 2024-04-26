using System.Data;
using System.IO.Compression;
using MiniExcelLibs;

namespace CsvToExcelConverter;

public class CsvToExcelConverter
{
    public static void ConvertCsvToExcel(string csvFilePath, int batchSize)
    {
        var lines = File.ReadLines(csvFilePath);
        var totalRecords = lines.Count();
        var numBatches = (int)Math.Ceiling((double)totalRecords / batchSize);

        // Read First Line as Column Name
        var columnNames = lines.First().Split(',').ToList();
        lines = lines.Skip(1);

        for (var i = 0; i < numBatches; i++)
        {
            Console.WriteLine($"Processing batch {i + 1} of {numBatches}...");

            var batchLines = lines.Skip(i * batchSize).Take(batchSize);
            var excelFilePath = Path.GetFileNameWithoutExtension(csvFilePath) + $"_{i + 1}.xlsx";
            
            var table = new DataTable();

            // Add Columns
            columnNames.ForEach(r => table.Columns.Add(r, typeof(string)));

            // Add Rows
            batchLines.ToList().ForEach(r =>
            {
                var row = table.NewRow();
                var values = r.Split(',');
                for(var j = 0; j < values.Length; j++)
                {
                    row[j] = values[j];
                }
                table.Rows.Add(row);
            });

            Console.WriteLine($"Saving batch {i + 1} to {excelFilePath}...");

            if (File.Exists(excelFilePath))
            {
                File.Delete(excelFilePath);
            }

            MiniExcel.SaveAs(excelFilePath, table);

            Console.WriteLine($"Batch {i + 1} saved to {excelFilePath}");

            table.Dispose();
        }

        // Compress all Excel files into a single ZIP file
        var zipFilePath = Path.GetFileNameWithoutExtension(csvFilePath) + ".zip";
        Console.WriteLine($"Compressing all Excel files into {zipFilePath}...");

        if (File.Exists(zipFilePath))
        {
            File.Delete(zipFilePath);
        }

        using (var zip = ZipFile.Open(zipFilePath, ZipArchiveMode.Create))
        {
            for(var i = 0; i < numBatches; i++)
            {
                var excelFilePath = Path.GetFileNameWithoutExtension(csvFilePath) + $"_{i + 1}.xlsx";
                zip.CreateEntryFromFile(excelFilePath, Path.GetFileName(excelFilePath));
                File.Delete(excelFilePath);
            }
        }

        Console.WriteLine($"All Excel files compressed into {zipFilePath}");
    }
}
