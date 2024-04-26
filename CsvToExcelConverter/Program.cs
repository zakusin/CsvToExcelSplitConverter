string? csvFileName = null;
string? batchSize = null;

if (Environment.GetCommandLineArgs().Length > 2)
{
    csvFileName = Environment.GetCommandLineArgs()[1];
    batchSize = Environment.GetCommandLineArgs()[2];
}

if (string.IsNullOrWhiteSpace(csvFileName))
{
    Console.WriteLine("Enter the CSV file name:");
    csvFileName = Console.ReadLine();
}

if (string.IsNullOrWhiteSpace(batchSize))
{
    Console.WriteLine("Enter the batch size:");
    batchSize = Console.ReadLine();
}

CsvToExcelConverter.CsvToExcelConverter.ConvertCsvToExcel($"{csvFileName}.csv", int.Parse(batchSize));

Console.WriteLine("Press any key to exit.");
Console.ReadLine();