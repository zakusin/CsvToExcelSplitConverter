using System.Data;
using System.IO.Compression;
using MiniExcelLibs;
using ClosedXML.Excel;

namespace CsvToExcelConverter;

public class CsvToExcelConverter
{
    public static void SplitExcelFileBatchedToXlsx(string xlsxFilePath, int batchSize)
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
