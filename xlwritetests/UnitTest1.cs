using System;
using NUnit.Framework;
using System.Text.RegularExpressions;
using xlwrite;

namespace Tests;

public class Tests
{
    [Test]
    public void CellReferenceParser()
    {
        string a1Reference = "B23";
        bool success  = XlWriteUtilities.TryParseCellReference(a1Reference, out Cell? cell);

        Assert.That(cell?.SheetName, Is.EqualTo(null));
        Assert.That(cell?.SheetNum, Is.EqualTo(-1));
        Assert.That(cell?.Column, Is.EqualTo(2));
        Assert.That(cell?.Row, Is.EqualTo(23));

        string r1c1Reference = "r12c65";
        success = XlWriteUtilities.TryParseCellReference(r1c1Reference, out cell);
        Assert.That(cell?.Column, Is.EqualTo(65));
        Assert.That(cell?.Row, Is.EqualTo(12));
    }

    [Test]
    public void NamedWorksheetTest()
    {
        string namedWorksheetReference = "'Sheet 1'!B23";
        bool success  = XlWriteUtilities.TryParseCellReference(namedWorksheetReference, out Cell? cell);
        Assert.That(cell?.SheetName, Is.EqualTo("Sheet 1"));
        Assert.That(cell?.SheetNum, Is.EqualTo(-1));
        Assert.That(cell?.Column, Is.EqualTo(2));
        Assert.That(cell?.Row, Is.EqualTo(23));
    }

    [Test]
    public void NumberedSheetTest()
    {
        string numberSheetReference = "1!B23";
        bool success  = XlWriteUtilities.TryParseCellReference(numberSheetReference, out Cell? cell);
        Assert.That(cell?.SheetName, Is.EqualTo(null));
        Assert.That(cell?.SheetNum, Is.EqualTo(1));
        Assert.That(cell?.Column, Is.EqualTo(2));
        Assert.That(cell?.Row, Is.EqualTo(23));
    }

    [Test]
    public void RegexTests()
    {
        string test = "]";
        Regex regex = new(@"[\]]");
        Assert.That(regex.Match(test).Success);
    }

    [Test]
    public void DateParseTest()
    {
        string test = "1/6";
        bool success = DateTime.TryParse(test, out DateTime dateTime);
        // Print ISO 8601 date format
        Console.WriteLine(dateTime.ToString("yyyy-MM-dd"));
        Assert.That(success);
    }
}