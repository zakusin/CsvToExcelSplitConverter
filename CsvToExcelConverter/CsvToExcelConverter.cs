using System.Data;
using System.IO.Compression;
using System.Xml.Linq;
using MiniExcelLibs;
using ClosedXML.Excel;

namespace CsvToExcelConverter;

public class CsvToExcelConverter
{
    public static void SplitCsvFileBatchedToXlsx(string csvFilePath, int batchSize)
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

            // Using ClosedXML to Open and set columns autofit
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1);

                // Autofit all columns
                worksheet.Columns().AdjustToContents();

                // Save the modified Excel file
                workbook.Save();
            }
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

    public static void SplitXlsxFileBatchedToXlsx(string xlsxFilePath, int batchSize)
    {
        var table = MiniExcel.QueryAsDataTable(xlsxFilePath);
        var totalRecords = table.Rows.Count;
        var numBatches = (int)Math.Ceiling((double)totalRecords / batchSize);

        for (var i = 0; i < numBatches; i++)
        {
            Console.WriteLine($"Processing batch {i + 1} of {numBatches}...");

            var batchTable = table.AsEnumerable().Skip(i * batchSize).Take(batchSize).CopyToDataTable();
            var batchExcelFilePath = Path.GetFileNameWithoutExtension(xlsxFilePath) + $"_{i + 1}.xlsx";

            Console.WriteLine($"Saving batch {i + 1} to {batchExcelFilePath}...");

            if (File.Exists(batchExcelFilePath))
            {
                File.Delete(batchExcelFilePath);
            }

            MiniExcel.SaveAs(batchExcelFilePath, batchTable);

            Console.WriteLine($"Batch {i + 1} saved to {batchExcelFilePath}");

            batchTable.Dispose();

            // Using ClosedXML to Open and set columns autofit
            using (var workbook = new XLWorkbook(batchExcelFilePath))
            {
                var worksheet = workbook.Worksheet(1);

                // Autofit all columns
                worksheet.Columns().AdjustToContents();

                // Save the modified Excel file
                workbook.Save();
            }
        }

        // Compress all Excel files into a single ZIP file
        var zipFilePath = Path.GetFileNameWithoutExtension(xlsxFilePath) + ".zip";
        Console.WriteLine($"Compressing all Excel files into {zipFilePath}...");

        if (File.Exists(zipFilePath))
        {
            File.Delete(zipFilePath);
        }

        using (var zip = ZipFile.Open(zipFilePath, ZipArchiveMode.Create))
        {
            for (var i = 0; i < numBatches; i++)
            {
                var excelFilePath = Path.GetFileNameWithoutExtension(xlsxFilePath) + $"_{i + 1}.xlsx";
                zip.CreateEntryFromFile(excelFilePath, Path.GetFileName(excelFilePath));
                File.Delete(excelFilePath);
            }
        }

        Console.WriteLine($"All Excel files compressed into {zipFilePath}");
    }
}
