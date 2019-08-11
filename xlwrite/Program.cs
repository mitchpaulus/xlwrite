using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace xlwrite
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("No arguments found. See help below.");
                Console.WriteLine(HelpText());
                return;
            }

            if (args.Any(arg => string.Equals("-h", arg) || string.Equals("--help", arg)))
            {
                Console.Write(HelpText());
                return;
            }

            string command = args[0];
            if (string.Equals(command, "block"))
            {
                if (args.Length < 4) {
                    Console.WriteLine("Not enough arguments for block.");
                    Console.WriteLine();
                    Console.WriteLine(HelpText());
                    return;
                }

                string blockResults = BlockWrite(args[1], args[2], args[3]);
                return;
            }

            else if (string.Equals(command, "ind"))
            {
                if (args.Length < 3) {
                    Console.WriteLine("Not enough arguments for ind.");
                    Console.WriteLine();
                    Console.WriteLine(HelpText());
                    return;
                }

                string indResults = IndWrite(args[1], args[2]);
                return;
            }
            else
            {
                Console.WriteLine($"Unknown sub command {command}. Please review help.");
                Console.WriteLine(HelpText());
                return;
            }
        }

        public static string BlockWrite(string cellReference, string dataFilename, string filename)
        {
            bool success = XlWriteUtilities.TryParseCellReference(cellReference, out Cell startCellLocation);
            if (!success) return $"Could not parse the cell reference {cellReference}.";

            var checkFiles = new List<string> { dataFilename, filename }.Select(s => new FileInfo(Path.Combine(Environment.CurrentDirectory, s))).ToList();
            if (checkFiles.Any(info => !info.Exists)) return $"Could not find file {checkFiles.First(info => !info.Exists)}.";

            try
            {
                var lines = File.ReadLines(checkFiles[0].FullName, Encoding.UTF8).Select(s => s.Split('\t'));

                List<(Cell cell, string value)> cells = new List<(Cell cell, string value)>();

                int index = 0;
                foreach (string[] fields in lines)
                {
                    int fieldIndex = 0;
                    foreach (string field in fields)
                    {
                        cells.Add((new Cell { Row = startCellLocation.Row + index, Column = startCellLocation.Column + fieldIndex }, field));
                        fieldIndex++;
                    }
                    index++;
                }

                ExcelPackage package = new ExcelPackage(checkFiles[1]);
                ExcelWorkbook workbook = package.Workbook;

                ExcelWorksheet sheet = XlWriteUtilities.SheetFromCell(package, startCellLocation);

                foreach ((Cell cell, string value) in cells)
                {
                    sheet.Cells[cell.Row, cell.Column].Value = value;
                }
                package.Save();
            }
            catch (Exception)
            {
                return $"There was an error with the parsing.";
            }

            return "";
        }

        public static string IndWrite(string dataFilename, string filename) {
            var checkFiles = new List<string> { dataFilename, filename }.Select(s => new FileInfo(Path.Combine(Environment.CurrentDirectory, s))).ToList();
            if (checkFiles.Any(info => !info.Exists)) return $"Could not find file {checkFiles.First(info => !info.Exists)}.";

            try {
                var lines = File.ReadLines(checkFiles[0].FullName, Encoding.UTF8).Select(s => s.Split('\t'));
                ExcelPackage package = new ExcelPackage(checkFiles[1]);

                int lineNumber = 1;
                foreach (var line in lines) {
                    if (line.Length != 2) {
                        Console.WriteLine($"Line #{lineNumber} does not have 2 fields. Found {line.Length} fields.");
                        continue;
                    }

                    bool success = XlWriteUtilities.TryParseCellReference(line[0], out Cell cell);
                    if (!success) {
                        Console.WriteLine($"Could not parse cell reference {line[0]}.");
                        continue;
                    }

                    ExcelWorksheet sheet = XlWriteUtilities.SheetFromCell(package, cell);

                    sheet.Cells[cell.Row, cell.Column].Value = line[1];

                    lineNumber++;
                }

                package.Save();
            }
            catch (Exception) {
                return "There was an error in the writing.";
            }

            return "";
        }

        public static string HelpText()
        {
            StringBuilder helpText = new StringBuilder();

            helpText.AppendLine("USAGE:");
            helpText.AppendLine("    xlwrite block STARTCELL DATAFILE EXCELFILE");
            helpText.AppendLine("    xlwrite ind DATAFILE EXCELFILE");
            helpText.AppendLine("    [-h | --help]");
            helpText.AppendLine();
            helpText.AppendLine("ARGS:");
            helpText.AppendLine($"    {"STARTCELL",-12}Upper left hand corner cell. Either A1 form or R1C1 form.");
            helpText.AppendLine($"    {"DATAFILE",-12}File with corresponding data.");
            helpText.AppendLine($"    {"EXCELFILE",-12}Excel file to insert data into.");
            helpText.AppendLine();
            helpText.AppendLine("OPTIONS:");
            helpText.AppendLine($"    {"-h, --help",-12}Print this help information and exit.");
            helpText.AppendLine();
            helpText.AppendLine("EXAMPLES:");
            helpText.AppendLine("    xlwrite block A1 mydata.tsv excelfile.xlsx");


            return helpText.ToString();
        }
    }

    public class Cell
    {
        public string SheetName;
        public int SheetNum;
        public int Row;
        public int Column;
    }

    public static class XlWriteUtilities
    {
        public static bool TryParseCellReference(string cellReference, out Cell cellLocation)
        {
            string worksheetNamePattern = @"'[^:\\/?*[\]]{1,31}'!";
            string worksheetNumberPattern = @"[1-9]\d+!";
            Regex a1Regex = new Regex($@"^(({worksheetNumberPattern})|({worksheetNamePattern}))?([A-Za-z]+)([0-9]+)$");
            Regex r1c1Regex = new Regex("^[rR]([0-9]+)[cC]([0-9]+)$");

            var a1Match = a1Regex.Match(cellReference);
            var r1c1Match = r1c1Regex.Match(cellReference);

            if (a1Match.Success)
            {
                cellLocation = new Cell
                {
                    SheetName = a1Match.Groups[3].Success ? a1Match.Groups[3].Value.Substring(1, a1Match.Groups[3].Value.Length - 3) : null,
                    SheetNum = a1Match.Groups[2].Success ? int.Parse(a1Match.Groups[2].Value.Substring(0, a1Match.Groups[2].Value.Length - 1)) : -1,
                    Row = int.Parse(a1Match.Groups[5].Value),
                    Column = a1Match.Groups[4].Value.ExcelColumnNameToInt()
                };
                return true;
            }
            else if (r1c1Match.Success)
            {
                cellLocation = new Cell
                {
                    Row = int.Parse(r1c1Match.Groups[1].Value),
                    Column = int.Parse(r1c1Match.Groups[2].Value)
                };
                return true;
            }
            else
            {
                cellLocation = null;
                return false;
            }
        }

        public static ExcelWorksheet SheetFromCell(ExcelPackage package, Cell cell) {
            ExcelWorkbook workbook = package.Workbook;

            if (workbook.Worksheets.Count == 0) {
                throw new InvalidOperationException($"There are no worksheets in file.");
            }

            if (cell.SheetNum > -1)
            {
                return workbook.Worksheets[cell.SheetNum - 1];
            }
            else if (cell.SheetName != null)
            {
                var matchingSheets = workbook.Worksheets.Where(s => s.Name == cell.SheetName).ToList();
                if (!matchingSheets.Any()) throw new InvalidOperationException($"Could not find sheet named {cell.SheetName}");
                return matchingSheets.First();
            }
            else
            {
                return workbook.Worksheets.First();
            }
        }
    }

    public static class StringExtensions
    {
        public static int ExcelColumnNameToInt(this string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException(nameof(columnName));

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            foreach (char c in columnName)
            {
                sum *= 26;
                sum += (c - 'A' + 1);
            }

            return sum;
        }
    }
}
