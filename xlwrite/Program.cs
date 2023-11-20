using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace xlwrite;

class Program
{
    static int Main(string[] args)
    {
        // This executable is free and open source, and is non-commercial.
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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
            Console.Write("0.6.0\n");
            return 0;
        }

        int argIndex = 0;
        bool createWorksheetIfRequired = false;
        bool autofitColumns = false;
        bool style = false;
        bool onePageWidth = false;
        bool onePageHeight = true;
        bool landscape = false;
        bool wipe = false;
        bool escape = false;
        bool debug = false;
        string? worksheet = null;

        while (argIndex < args.Length)
        {
            var arg = args[argIndex];
            if (arg is "-c" or "--create")
            {
                createWorksheetIfRequired = true; argIndex++;
            }
            else if (args[argIndex] == "-a" || args[argIndex] == "--autofit" || args[argIndex] == "--auto-fit")
            {
                autofitColumns = true; argIndex++;
            }
            else if (args[argIndex] is "--debug")
            {
                debug = true; argIndex++;
            }
            else if (args[argIndex] == "--style")
            {
                style = true; argIndex++;
            }
            else if (args[argIndex] is "-e" or "--escape")
            {
                escape = true; argIndex++;
            }
            else if (arg is "-w" or "--worksheet" or "--sheet")
            {
                if (argIndex + 1 >= args.Length)
                {
                    Console.Error.Write($"No sheet given to option {arg}.\n\n");
                    return 1;
                }
                worksheet = args[argIndex + 1];
                argIndex += 2;
            }
            else if (arg is "--1-page-width" or "--1pagewidth")
            {
                onePageWidth = true; argIndex++;
            }
            else if (arg is "--1-page-height" or "--1pageheight")
            {
                onePageHeight = true; argIndex++;
            }
            else if (arg is "--landscape")
            {
                landscape = true; argIndex++;
            }
            else if (arg is "--wipe")
            {
                wipe = true; argIndex++;
            }
            else
            {
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


                    int remainingArgCount = args.Length - argIndex;

                    List<(string CellReference, string DataFilename, string? Worksheet)> cellDataFilenames = new();
                    string excelFilename;

                    if (args.Length - argIndex == 4)
                    {
                        string cellReference = args[argIndex + 1];
                        string dataFilename = args[argIndex + 2];
                        excelFilename = args[argIndex + 3];

                        cellDataFilenames.Add((cellReference, dataFilename, worksheet));
                    }
                    else if (remainingArgCount < 4)
                    {
                        Console.Error.Write($"Not enough arguments passed after block command\n. There should be a minimum of 3 (datafile, cell, Excel file), received {remainingArgCount - 1}\n");
                        return 1;
                    }
                    else if ((remainingArgCount - 2) % 3 == 0)
                    {
                        excelFilename = args[argIndex + 1];

                        // Read remaining arguments in threes.
                        for (int i = argIndex + 2; i < args.Length; i += 3)
                        {
                            string worksheetArg = args[i];
                            string cellReference = args[i + 1];
                            string dataFilename = args[i + 2];

                            cellDataFilenames.Add((cellReference, dataFilename, worksheetArg));
                        }
                    }
                    else
                    {
                        Console.Error.Write($"Expected a multiple of 3 number of arguments after the 'block' command and Excel file. Got {remainingArgCount - 2}.\n");
                        return 1;
                    }

                    ExcelPackage package;
                    try
                    {
                        FileInfo excelFile = new(Path.Combine(Environment.CurrentDirectory, excelFilename));

                        if (wipe)
                        {
                            try
                            {
                                File.Delete(excelFile.FullName);
                            }
                            catch
                            {
                                Console.Error.Write($"Could not delete file '{excelFile.FullName}'");
                                return 1;
                            }
                        }

                        package = new(excelFile);
                    }
                    catch
                    {
                        Console.Error.Write($"Could not open Excel file: {excelFilename}\n");
                        return 1;
                    }

                    // Loop over the cell and data file pairs.
                    // The reason for this level of effort is that the actual saving of the file takes significant time,
                    // so instead of looping over input files and saving on each individual write, you can specify them
                    // all from the command line in one command, resulting in a single save.
                    foreach ((string cellRef, string dataFile, string? sheet) in cellDataFilenames)
                    {
                        if (debug) Console.Error.Write($"Writing '{dataFile}' to '{sheet ?? ""}': {cellRef}\n");
                        string blockResults = BlockWrite(cellRef, dataFile, package, createWorksheetIfRequired, autofitColumns, style, sheet, wipe, escape, debug);
                        if (string.IsNullOrWhiteSpace(blockResults)) continue;

                        Console.Error.Write(blockResults);
                        return 1;
                    }

                    // Only save once at the end.
                    if (debug)
                    {
                        Stopwatch watch = new();
                        watch.Start();
                        package.Save();
                        watch.Stop();
                        Console.Error.Write($"Saving file: {watch.ElapsedMilliseconds}ms\n");
                    }
                    else
                    {
                        package.Save();
                    }

                    SetPageWidth(excelFilename, onePageWidth, onePageHeight, landscape);
                    return 0;
                }

                if (string.Equals(command, "ind"))
                {
                    if (args.Length - argIndex < 3)
                    {
                        Console.Error.Write("Not enough arguments for ind.\n\n");
                        Console.Error.Write(HelpText());
                        return 1;
                    }

                    string dataFilename = args[argIndex + 1];
                    string excelFilename = args[argIndex + 2];
                    string indResults = IndWrite(dataFilename, excelFilename, createWorksheetIfRequired, wipe, escape);

                    SetPageWidth(excelFilename, onePageWidth, onePageHeight, landscape);

                    if (string.IsNullOrWhiteSpace(indResults)) return 0;
                    Console.Error.Write(indResults.EndWithNewline());
                    return 1;
                }

                Console.Error.Write($"Unknown sub command {command}. Please review help.\n");
                Console.Error.Write(HelpText());
                return 1;
            }
        }

        // Shouldn't get here.
        return 0;
    }

    public static void SetPageWidth(string filePath, bool fitToWidth, bool fitToHeight, bool landscape)
    {
        if (!fitToHeight && !fitToHeight) return;

        ExcelPackage package = new ExcelPackage(filePath);
        foreach (var sheet in package.Workbook.Worksheets)
        {
            sheet.PrinterSettings.FitToPage = true;
            if (fitToWidth) sheet.PrinterSettings.FitToWidth = 1;
            if (fitToHeight) sheet.PrinterSettings.FitToHeight = 1;
            if (landscape) sheet.PrinterSettings.Orientation = eOrientation.Landscape;

        }
        package.Save();
    }

    public static string BlockWrite(string cellReference, string dataFilename, ExcelPackage package, bool createWorksheetIfRequired, bool autoFitColumns, bool style, string? worksheet, bool wipe, bool escape, bool debug)
    {
        Stopwatch watch = new();
        if (!XlWriteUtilities.TryParseCellReference(cellReference, out Cell? startCellLocation)) return $"Could not parse the cell reference {cellReference}.";

        List<FileInfo> checkFiles = new List<string> { dataFilename }
            .Where(name => !string.Equals("-", name))
            .Select(s => new FileInfo(Path.Combine(Environment.CurrentDirectory, s)))
            .ToList();

        if (checkFiles.Any(info => !info.Exists)) return $"Could not find file {checkFiles.First(info => !info.Exists)}.";

        watch.Restart();
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
                FileInfo fullDataFilename = new(Path.Combine(Environment.CurrentDirectory, dataFilename));
                li = File.ReadLines(fullDataFilename.FullName, Encoding.UTF8).ToList();
            }
        }
        catch (Exception)
        {
            return "Could not read data from data source.";
        }


        IEnumerable<string[]> lines = li.Select(s => s.Split('\t'));
        List<(Cell cell, string value)> cells = new();

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
        watch.Stop();

        if (debug) Console.Error.Write($"Read data from source in {watch.ElapsedMilliseconds}ms\n");

        try
        {
            ExcelWorksheet sheet = string.IsNullOrWhiteSpace(worksheet)
                ? XlWriteUtilities.SheetFromCell(package, startCellLocation, createWorksheetIfRequired)
                : XlWriteUtilities.SheetFromName(package, worksheet, createWorksheetIfRequired);

            watch.Restart();
            HashSet<int> columnsUsed = new();
            if (escape)
            {
                foreach ((Cell cell, string value) in cells)
                {
                    sheet.Cells[cell.Row, cell.Column].Value = GetEscapedValue(value);
                    columnsUsed.Add(cell.Column);
                }
            }
            else
            {
                foreach ((Cell cell, string value) in cells)
                {
                    sheet.Cells[cell.Row, cell.Column].Value = GetValue(value);
                    columnsUsed.Add(cell.Column);
                }
            }
            watch.Stop();
            if (debug) Console.Error.Write($"Pushed {cells.Count} cells in {watch.ElapsedMilliseconds}ms\n");

            if (style)
            {
                if (cells.Any())
                {
                    int headerRow = cells.Min(c => c.cell.Row);
                    int startColumn = cells.Min(c => c.cell.Column);
                    int endRow = cells.Max(c => c.cell.Row);
                    int endColumn = cells.Max(c => c.cell.Column);

                    for (int row = headerRow; row <= endRow; row++)
                    {
                        for (int column = startColumn; column <= endColumn; column++)
                        {
                            sheet.Cells[row, column].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                    }

                    // Loop over header cells, make them bold, white text, CCLLC blue background
                    for (int column = startColumn; column <= endColumn; column++)
                    {
                        ExcelStyle? excelStyle = sheet.Cells[headerRow, column].Style;
                        excelStyle.Font.Bold = true;
                        excelStyle.Font.Color.SetColor(Color.White);
                        excelStyle.Fill.PatternType = ExcelFillStyle.Solid;
                        excelStyle.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 73, 135));
                        excelStyle.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                }

                sheet.PrinterSettings.FitToPage = true;
                sheet.PrinterSettings.FitToWidth = 1;
                sheet.PrinterSettings.FitToHeight = 0;
                sheet.PrinterSettings.Orientation = eOrientation.Landscape;
                sheet.HeaderFooter.OddFooter.CenteredText = sheet.Name;
                sheet.HeaderFooter.ScaleWithDocument = false;
            }

            if (autoFitColumns) foreach (int colNum in columnsUsed) sheet.Column(colNum).AutoFit();
        }
        catch (Exception exception)
        {
            string errorMessage = $"There was an error with writing the data to the excel file.\n{exception.Message}\n";
            if (exception.InnerException != null) errorMessage += exception.InnerException.Message;

            return errorMessage;
        }

        return "";
    }

    public static string IndWrite(string dataFilename, string filename, bool createWorksheetIfRequired, bool wipe, bool escape)
    {
        List<FileInfo> checkFiles = new List<string> { dataFilename, filename }.Select(s => new FileInfo(Path.Combine(Environment.CurrentDirectory, s))).ToList();
        if (checkFiles.Any(info => !info.Exists)) return $"Could not find file {checkFiles.First(info => !info.Exists)}.";

        try
        {
            IEnumerable<string[]> lines = File.ReadLines(checkFiles[0].FullName, Encoding.UTF8).Select(s => s.Split('\t'));


            FileInfo excelFile = checkFiles[1];
            if (wipe)
            {
                try { File.Delete(excelFile.FullName); }
                catch { return $"Could not delete file '{excelFile.FullName}'"; }
            }

            ExcelPackage package = new(excelFile);

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

                sheet.Cells[cell.Row, cell.Column].Value = escape ? GetEscapedValue(field[1]) : GetValue(field[1]);

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

    public static object GetEscapedValue(string data)
    {
        if (double.TryParse(data, out double numericValue)) return numericValue;
        // This is to prevent fractions like 1/6 from being converted to dates.
        if (data.Length > 8 && DateTime.TryParse(data, out DateTime dateTime)) return dateTime;
        return data.ProcessEscapeSequences();
    }

    public static string HelpText()
    {
        StringBuilder helpText = new();

        const int padding = -12;
        const int optionPadding = -15;

        // ReSharper disable StringLiteralTypo
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
        helpText.AppendLine($"    {"-a, --autofit",optionPadding}Auto fit columns for which data has been entered. (Only 'block' mode currently).");
        helpText.AppendLine($"    {"-c, --create",optionPadding}Create specified worksheet if required.");
        helpText.AppendLine($"    {"-e, --escape",optionPadding}Process escape sequences '\\n' and '\\t'");
        helpText.AppendLine($"    {"-h, --help",optionPadding}Print this help information and exit.");
        helpText.AppendLine($"    {"-w, --sheet",optionPadding}Specify worksheet. Only affect block mode.");
        helpText.AppendLine($"    {"    --wipe",optionPadding}Delete existing file before writing. Be careful!");
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
        // ReSharper restore StringLiteralTypo


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
        Regex a1Regex = new($@"^(({worksheetNumberPattern})|({worksheetNamePattern}))?([A-Za-z]+)([0-9]+)$");
        Regex rowColRegex = new("^[rR]([0-9]+)[cC]([0-9]+)$");

        Match a1Match = a1Regex.Match(cellReference);
        Match rowColMatch = rowColRegex.Match(cellReference);

        if (a1Match.Success)
        {
            cellLocation = new Cell
            {
                // The 1 is to skip the first quote, and the -2 on the length is for the quote and exclamation point.
                SheetName = a1Match.Groups[3].Success ? a1Match.Groups[3].Value[1..^2] : null,
                // The ^1 is to remove the ! at the end of the pattern.
                SheetNum = a1Match.Groups[2].Success ? int.Parse(a1Match.Groups[2].Value[..^1]) : -1,
                Row = int.Parse(a1Match.Groups[5].Value),
                Column = a1Match.Groups[4].Value.ExcelColumnNameToInt()
            };
            return true;
        }

        if (rowColMatch.Success)
        {
            cellLocation = new Cell
            {
                Row = int.Parse(rowColMatch.Groups[1].Value),
                Column = int.Parse(rowColMatch.Groups[2].Value)
            };
            return true;
        }

        cellLocation = null;
        return false;
    }

    public static ExcelWorksheet SheetFromCell(ExcelPackage package, Cell cell, bool createSheetIfRequired)
    {
        ExcelWorksheets sheets = package.Workbook.Worksheets;
        bool sheetSpecified = cell.SheetNum > 0 || cell.SheetName != null;

        if (!sheetSpecified) return sheets.Any() ? sheets.First() : sheets.Add("Sheet 1");
        if (cell.SheetName != null) return SheetFromName(package, cell.SheetName, createSheetIfRequired);
        if (cell.SheetNum <= sheets.Count) return sheets[cell.SheetNum - 1];

        throw new InvalidOperationException($"Specified sheet index {cell.SheetNum}' but only {package.Workbook.Worksheets.Count} sheets exist in file '{package.File.FullName}'.");
    }

    public static ExcelWorksheet SheetFromName(ExcelPackage package, string name, bool createSheetIfRequired)
    {
        ExcelWorksheets sheets = package.Workbook.Worksheets;
        ExcelWorksheet? matchingSheet = sheets.FirstOrDefault(s => s.Name == name);
        if (matchingSheet is not null) return matchingSheet;
        if (createSheetIfRequired)
        {
            return sheets.Add(name.SanitizeExcelSheetName());
        }
        throw new InvalidOperationException($"Could not find sheet named '{name}' in file '{package.File.FullName}'.");

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

    public static string SanitizeExcelSheetName(this string inputSheetName)
    {
        //Replace invalid characters. https://www.accountingweb.com/technology/excel/seven-characters-you-cant-use-in-worksheet-names
        inputSheetName = inputSheetName.Replace("/", " ");
        inputSheetName = inputSheetName.Replace("\\", " ");
        inputSheetName = inputSheetName.Replace("?", " ");
        inputSheetName = inputSheetName.Replace("*", " ");
        inputSheetName = inputSheetName.Replace("[", " ");
        inputSheetName = inputSheetName.Replace("]", " ");
        inputSheetName = inputSheetName.Replace(":", " ");

        // Excel sheet names cannot be the word "history".
        if (inputSheetName.ToLower() == "history")
        {
            Random random = new Random();
            inputSheetName += "_" + random.Next(1, 1000);
        }

        return inputSheetName.Truncate(31, "..");  // Excel has a limit of 31 characters in sheet names. https://stackoverflow.com/questions/3681868/is-there-a-limit-on-an-excel-worksheets-name-length
    }

    /// <summary>
    /// Truncates string to the first <paramref name="maxLength"/> characters. If characters are removed, then the <paramref name="append"/> string is appended.
    /// </summary>
    /// <param name="value"></param>
    /// <param name="maxLength"></param>
    /// <param name="append">What to append if characters are removed. Example: "..."</param>
    /// <returns></returns>
    public static string Truncate(this string value, int maxLength, string append = "")
    {
        if (append.Length > maxLength)
            throw new Exception("Append string must not be greater than maxLength.");

        // Pulled from https://stackoverflow.com/questions/2776673/how-do-i-truncate-a-net-string/2776720
        if (string.IsNullOrEmpty(value)) return value;
        return value.Length <= maxLength
            ? value
            : value[..(maxLength - append.Length)] + append;
    }

    public static string ProcessEscapeSequences(this string input)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));

        StringBuilder result = new(input.Length + 100);
        bool escapeNext = false;

        foreach (char c in input)
        {
            if (escapeNext)
            {
                switch (c)
                {
                    case 'n':
                        result.Append('\n');
                        break;
                    case '\t':
                        result.Append('\t');
                        break;
                    case '\\':
                        result.Append('\\');
                        break;
                    default:
                        result.Append('\\');
                        result.Append(c);
                        break;
                }
                escapeNext = false;
            }
            else if (c == '\\')
            {
                escapeNext = true;
            }
            else
            {
                result.Append(c);
            }
        }

        if (escapeNext) result.Append('\\');
        return result.ToString();
    }
}
