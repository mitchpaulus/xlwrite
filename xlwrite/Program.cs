using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace xlwrite
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.Error.Write("No arguments found. See help below.\n");
                Console.Error.Write(HelpText());
                return 1;
            }

            if (args.Any(s => s is "-h" or "--help"))
            {
                Console.Write(HelpText());
                return 0;
            }

            if (args.Any(s => s is "-v" or "--version"))
            {
                Console.Write("0.3.0\n");
                return 0;
            }

            int argIndex = 0;
            bool createWorksheetIfRequired = false;

            if (args[argIndex] == "-c" || args[argIndex] == "--create")
            {
                createWorksheetIfRequired = true;
                argIndex++;
            }

            string command = args[argIndex];
            if (string.Equals(command, "block"))
            {
                if (args.Length - argIndex < 4)
                {
                    Console.Error.Write("Not enough arguments for block.\n");
                    Console.Error.Write("\n");
                    Console.Error.Write(HelpText());
                    return 1;
                }

                string blockResults = BlockWrite(args[argIndex + 1], args[argIndex + 2], args[argIndex + 3], createWorksheetIfRequired);
                if (string.IsNullOrWhiteSpace(blockResults)) return 0;
                Console.Error.WriteLine(blockResults);
                return 1;

            }

            if (string.Equals(command, "ind"))
            {
                if (args.Length - argIndex < 3)
                {
                    Console.Error.Write("Not enough arguments for ind.\n\n");
                    Console.Error.Write(HelpText());
                    return 1;
                }

                string indResults = IndWrite(args[argIndex + 1], args[argIndex + 2], createWorksheetIfRequired);
                if (string.IsNullOrWhiteSpace(indResults)) return 0;
                Console.Error.Write(indResults.EndWithNewline());
                return 1;
            }

            Console.Error.WriteLine($"Unknown sub command {command}. Please review help.\n");
            Console.Error.WriteLine(HelpText());
            return 1;
        }

        public static string BlockWrite(string cellReference, string dataFilename, string filename, bool createWorksheetIfRequired)
        {
            if (!XlWriteUtilities.TryParseCellReference(cellReference, out Cell? startCellLocation)) return $"Could not parse the cell reference {cellReference}.";

            List<FileInfo> checkFiles = new List<string> { dataFilename }
                .Where(name => !string.Equals("-", name))
                .Select(s => new FileInfo(Path.Combine(Environment.CurrentDirectory, s)))
                .ToList();

            if (checkFiles.Any(info => !info.Exists)) return $"Could not find file {checkFiles.First(info => !info.Exists)}.";

            List<string> li;
            try
            {
                if (dataFilename == "-")
                {
                    li = new List<string>();
                    using TextReader reader = Console.In;
                    while (reader.ReadLine() is { } text)
                    {
                        li.Add(text);
                    }
                }
                else
                {
                    FileInfo fullDataFilename = new FileInfo(Path.Combine(Environment.CurrentDirectory, dataFilename));
                    li = File.ReadLines(fullDataFilename.FullName, Encoding.UTF8).ToList();
                }
            }
            catch (Exception)
            {
                return "Could not read data from data source.";
            }


            IEnumerable<string[]> lines = li.Select(s => s.Split('\t'));
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
            try
            {
                FileInfo excelFile = new FileInfo(Path.Combine(Environment.CurrentDirectory, filename));
                ExcelPackage package = new ExcelPackage(excelFile);

                // if (!excelFile.Exists) package.Workbook.Worksheets.Add("Sheet 1");

                ExcelWorksheet sheet = XlWriteUtilities.SheetFromCell(package, startCellLocation, createWorksheetIfRequired);

                foreach ((Cell cell, string value) in cells)
                {
                    sheet.Cells[cell.Row, cell.Column].Value = GetValue(value);
                }
                package.Save();
            }
            catch (Exception exception)
            {
                string errorMessage = $"There was an error with writing the data to the excel file.\n{exception.Message}\n";
                if (exception.InnerException != null) errorMessage += exception.InnerException.Message;

                return errorMessage;
            }

            return "";
        }

        public static string IndWrite(string dataFilename, string filename, bool createWorksheetIfRequired)
        {
            var checkFiles = new List<string> { dataFilename, filename }.Select(s => new FileInfo(Path.Combine(Environment.CurrentDirectory, s))).ToList();
            if (checkFiles.Any(info => !info.Exists)) return $"Could not find file {checkFiles.First(info => !info.Exists)}.";

            try
            {
                IEnumerable<string[]> lines = File.ReadLines(checkFiles[0].FullName, Encoding.UTF8).Select(s => s.Split('\t'));

                ExcelPackage package = new ExcelPackage(checkFiles[1]);
                int lineNumber = 1;
                foreach (string[] field in lines)
                {
                    if (field.Length != 2)
                    {
                        Console.Error.Write($"Line #{lineNumber} does not have 2 fields. Found {field.Length} fields.\n");
                        continue;
                    }

                    if (!XlWriteUtilities.TryParseCellReference(field[0], out Cell? cell))
                    {
                        Console.Error.Write($"Could not parse cell reference {field[0]}.\n");
                        continue;
                    }

                    ExcelWorksheet sheet = XlWriteUtilities.SheetFromCell(package, cell, createWorksheetIfRequired);

                    sheet.Cells[cell.Row, cell.Column].Value = GetValue(field[1]);

                    lineNumber++;
                }

                package.Save();
            }
            catch (Exception)
            {
                return "There was an error in the writing.";
            }

            return "";
        }

        public static object GetValue(string data)
        {
            if (double.TryParse(data, out double numericValue)) return numericValue;
            // This is to prevent fractions like 1/6 from being converted to dates.
            if (data.Length > 8 && DateTime.TryParse(data, out DateTime dateTime)) return dateTime;
            return data;
        }

        public static string HelpText()
        {
            StringBuilder helpText = new StringBuilder();

            const int padding = -12;
            const int optionPadding = -15;

            helpText.AppendLine("USAGE:");
            helpText.AppendLine("    xlwrite [OPTION].. block STARTCELL DATAFILE EXCELFILE");
            helpText.AppendLine("    xlwrite [OPTION].. ind DATAFILE EXCELFILE");
            helpText.AppendLine();
            helpText.AppendLine("ARGS:");
            helpText.AppendLine($"    {"STARTCELL",padding}Upper left hand corner cell. Either A1 form or R1C1 form.");
            helpText.AppendLine($"    {"DATAFILE",padding}File with corresponding data. '-' can be used to read in standard input.");
            helpText.AppendLine($"    {"EXCELFILE",padding}Excel file to insert data into. If file doesn't exist, a new file is created.");
            helpText.AppendLine();
            helpText.AppendLine("OPTIONS:");
            helpText.AppendLine($"    {"-c, --create",optionPadding}Create specified worksheet if required.");
            helpText.AppendLine($"    {"-h, --help",optionPadding}Print this help information and exit.");
            helpText.AppendLine($"    {"-v, --version",optionPadding}Print version information and exit.");
            helpText.AppendLine();
            helpText.AppendLine("EXAMPLES:");
            helpText.AppendLine("Simple block usage:");
            helpText.AppendLine("    xlwrite block A1 mydata.tsv excelfile.xlsx");
            helpText.AppendLine();
            helpText.AppendLine("Reading from standard input:");
            helpText.AppendLine("    cat myfile.tsv | xlwrite block A1 - excelfile.xlsx");
            helpText.AppendLine();
            helpText.AppendLine("If you are using the 'ind' option, the format of the file is:");
            helpText.AppendLine();
            helpText.AppendLine("<Cell Reference><Tab><Value>");
            helpText.AppendLine();
            helpText.AppendLine("For example:");
            helpText.AppendLine();
            helpText.AppendLine("B1\t101");
            helpText.AppendLine("E5\tsome text");
            helpText.AppendLine("");
            helpText.AppendLine("If you want to specify a cell on a particular sheet, you can use the Excel format.");
            helpText.AppendLine("Note that you will likely need some shell quoting to get the apostrophes through.");
            helpText.AppendLine();
            helpText.AppendLine("xlwrite block \"'Sheet 2'!B2\" data.txt excel.xlsx");


            return helpText.ToString();
        }
    }

    public class Cell
    {
        public string? SheetName;
        /// <summary>
        /// 1-based Worksheet number index
        /// </summary>
        public int SheetNum;
        public int Row;
        public int Column;
    }

    public static class XlWriteUtilities
    {
        public static bool TryParseCellReference(string cellReference, [NotNullWhen(returnValue:true)] out Cell? cellLocation)
        {
            string worksheetNamePattern = @"'[^:\\/?*[\]]{1,31}'!";
            string worksheetNumberPattern = @"[1-9]\d*!";
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

        public static ExcelWorksheet SheetFromCell(ExcelPackage package, Cell cell, bool createSheetIfRequired)
        {
            ExcelWorksheets sheets = package.Workbook.Worksheets;
            bool sheetSpecified = cell.SheetNum > 0 || cell.SheetName != null;
            if (sheetSpecified)
            {
                if (cell.SheetName != null)
                {
                    List<ExcelWorksheet> matchingSheets = sheets.Where(s => s.Name == cell.SheetName).ToList();
                    if (matchingSheets.Any()) return matchingSheets.First();
                    if (createSheetIfRequired) return sheets.Add(cell.SheetName);
                    throw new InvalidOperationException($"Could not find sheet named '{cell.SheetName}' in file '{package.File.FullName}'.");
                }

                if (cell.SheetNum <= sheets.Count) return sheets[cell.SheetNum - 1];

                throw new InvalidOperationException($"Specified sheet index {cell.SheetNum}' but only {package.Workbook.Worksheets.Count} sheets exist in file '{package.File.FullName}'.");
            }

            return sheets.Any() ? sheets.First() : sheets.Add("Sheet 1");
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

        // Check if string ends with Unix newline, if so, return it, else add newline
        public static string EndWithNewline(this string inputString) => inputString.EndsWith("\n") ? inputString : inputString + "\n";
    }
}
