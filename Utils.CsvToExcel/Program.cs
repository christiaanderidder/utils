using System.CommandLine;
using System.Globalization;
using ClosedXML.Excel;
using CsvHelper;

var rootCommand = new RootCommand("CSV to Excel converter");
var inputDirOption = new Option<DirectoryInfo>("--inputDir", "The directory containing the CSV files to convert");
var outputDirOption = new Option<DirectoryInfo>("--outputDir", "The where the resulting Excel files will be saved");
rootCommand.Add(inputDirOption);
rootCommand.Add(outputDirOption);

rootCommand.SetHandler((inputDir, outputDir) =>
{
    var files = inputDir.EnumerateFiles("*.csv");
    Parallel.ForEach(files, file => ProcessFile(file, outputDir));
}, inputDirOption, outputDirOption);

await rootCommand.InvokeAsync(args);

static void ProcessFile(FileInfo file, DirectoryInfo outputDir)
{
    using var reader = new StreamReader(file.FullName);
    using var csv = new CsvParser(reader, CultureInfo.InvariantCulture);
    
    var outputPath = Path.Combine(outputDir.FullName, $"{Path.GetFileNameWithoutExtension(file.Name)}.xlsx");

    using var workbook = new XLWorkbook();
    var worksheet = workbook.Worksheets.Add("Sheet 1");

    var data = new List<string[]?>();
    while (csv.Read())
    {
        data.Add(csv.Record);
    }

    worksheet.Cell(1, 1).InsertData(data);
    worksheet.Cell(1, 1).WorksheetRow().Style.Font.Bold = true;
    
    using var outputFile = File.Open(outputPath, FileMode.OpenOrCreate);
    workbook.SaveAs(outputFile);
}