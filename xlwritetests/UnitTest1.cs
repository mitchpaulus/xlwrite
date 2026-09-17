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

    [Test]
    public void ColumnListParser()
    {
        Assert.That(XlWriteUtilities.TryParseColumnList("A,C", out var cols), Is.True);
        Assert.That(cols, Is.EquivalentTo(new[] { 1, 3 }));

        Assert.That(XlWriteUtilities.TryParseColumnList("1, 28,ab", out cols), Is.True);
        Assert.That(cols, Is.EquivalentTo(new[] { 1, 28 }));

        Assert.That(XlWriteUtilities.TryParseColumnList("A1", out _), Is.False);
        Assert.That(XlWriteUtilities.TryParseColumnList("0", out _), Is.False);
        Assert.That(XlWriteUtilities.TryParseColumnList("", out _), Is.False);
    }

    [Test]
    public void OutOfRangeNumberStaysText()
    {
        Assert.That(Program.GetValue("9E750"), Is.EqualTo("9E750"));
        Assert.That(Program.GetEscapedValue("9E750"), Is.EqualTo("9E750"));
        Assert.That(Program.GetValue("Infinity"), Is.EqualTo("Infinity"));
        Assert.That(Program.GetValue("9E5"), Is.EqualTo(900000d));
        Assert.That(Program.GetValue("9E5", forceText: true), Is.EqualTo("9E5"));
        Assert.That(Program.GetValue("2024-01-05", forceText: true), Is.EqualTo("2024-01-05"));
    }
}