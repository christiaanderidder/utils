﻿using System.CommandLine;
using System.Globalization;
using ClosedXML.Excel;
using CsvHelper;

namespace Utils.CsvToExcel;

public static class Program
{
    public static async Task Main(string[] args)
    {
        var rootCommand = new RootCommand("CSV to Excel converter");
        var inputDirOption = new Option<DirectoryInfo>("--inputDir", "The directory containing the CSV files to convert") {IsRequired = true};
        var outputDirOption = new Option<DirectoryInfo>("--outputDir", "The directory where the resulting Excel files will be saved") {IsRequired = true};
        rootCommand.Add(inputDirOption);
        rootCommand.Add(outputDirOption);

        rootCommand.SetHandler((inputDir, outputDir) =>
        {
            var files = inputDir.EnumerateFiles("*.csv").ToList();
            Console.WriteLine($"Found {files.Count} files for pattern *.csv in {outputDir.FullName}");
            
            Parallel.ForEach(files, file => ProcessFile(file, outputDir));
            
            Console.WriteLine("Done");
        }, inputDirOption, outputDirOption);

        await rootCommand.InvokeAsync(args);
    }

    private static void ProcessFile(FileInfo file, DirectoryInfo outputDir)
    {
        using var reader = new StreamReader(file.FullName);
        using var csv = new CsvParser(reader, CultureInfo.InvariantCulture);

        var filename = Path.GetFileNameWithoutExtension(file.Name);
        var outputPath = Path.Combine(outputDir.FullName, $"{filename}.xlsx");

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet 1");

        Console.WriteLine($"[{filename}] Reading CSV");
        var data = new List<string[]?>();
        while (csv.Read())
        {
            data.Add(csv.Record);
        }

        Console.WriteLine($"[{filename}] Writing Excel");
        worksheet.Cell(1, 1).InsertData(data);
        worksheet.Cell(1, 1).WorksheetRow().Style.Font.Bold = true;

        Console.WriteLine($"[{filename}] Saving Excel");
        using var outputFile = File.Open(outputPath, FileMode.OpenOrCreate);
        workbook.SaveAs(outputFile);
        
        Console.WriteLine($"[{filename}] Done");
    }
}