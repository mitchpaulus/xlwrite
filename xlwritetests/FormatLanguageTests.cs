using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using xlwrite;

namespace Tests;

public class FormatLanguageTests
{
    private static FormatScript Parse(string source)
    {
        (FormatScript? script, List<string> errors) = FormatCompiler.Compile(source);
        Assert.That(errors, Is.Empty);
        return script!;
    }

    private static List<string> Errors(string source)
    {
        (FormatScript? script, List<string> errors) = FormatCompiler.Compile(source);
        Assert.That(script, Is.Null);
        return errors;
    }

    [Test]
    public void ItemsNeedNoSeparators()
    {
        FormatScript script = Parse("A1:D1 bold fill #DDEBF7 B:B width 20 2:3 h 30");
        Assert.That(script.Items, Has.Count.EqualTo(3));
        Assert.That(script.Items[0].Elements, Is.EqualTo(new FormatElement[]
        {
            new ToggleAction(ToggleProperty.Bold, true),
            new FillAction(new Rgb(0xDD, 0xEB, 0xF7)),
        }));
        Assert.That(script.Items[1].Selections.Single().Kind, Is.EqualTo(RefKind.Columns));
        Assert.That(script.Items[2].Selections.Single().Kind, Is.EqualTo(RefKind.Rows));
    }

    [Test]
    public void SheetFormsAndStickiness()
    {
        FormatScript script = Parse("'It''s'!A1 A2 \"Other\" B1 Data!C1 C2 bold D1 italic");
        List<Selection> s = script.Items[0].Selections;
        Assert.That(s.Select(x => x.Sheet), Is.EqualTo(new[] { "It's", "It's", "Other", "Data", "Data" }));
        Assert.That(script.Items[1].Selections.Single().Sheet, Is.Null);
    }

    [Test]
    public void TableSelections()
    {
        FormatScript script = Parse("a1:* bold B2:* header bold B2:* col \"Total\" fmt \"0.00\" B2:* body bold B2:* lastrow bold");
        Assert.That(script.Items[0].Selections.Single(), Is.EqualTo(new Selection(null, RefKind.Table, "A1")));
        Assert.That(script.Items[1].Selections.Single().Part, Is.EqualTo(TablePart.Header));
        Assert.That(script.Items[2].Selections.Single(), Is.EqualTo(new Selection(null, RefKind.Table, "B2", TablePart.Column, "Total")));
        Assert.That(script.Items[3].Selections.Single().Part, Is.EqualTo(TablePart.Body));
        Assert.That(script.Items[4].Selections.Single().Part, Is.EqualTo(TablePart.LastRow));
    }

    [Test]
    public void TogglesAndColors()
    {
        FormatScript script = Parse("A1 bold off italic on underline true strike false wrap fill red color rgb 1 2 3 bg #0a0B0c");
        Assert.That(script.Items[0].Elements, Is.EqualTo(new FormatElement[]
        {
            new ToggleAction(ToggleProperty.Bold, false),
            new ToggleAction(ToggleProperty.Italic, true),
            new ToggleAction(ToggleProperty.Underline, true),
            new ToggleAction(ToggleProperty.Strike, false),
            new ToggleAction(ToggleProperty.Wrap, true),
            new FillAction(new Rgb(255, 0, 0)),
            new FontColorAction(new Rgb(1, 2, 3)),
            new FillAction(new Rgb(10, 11, 12)),
        }));
    }

    [Test]
    public void Borders()
    {
        FormatScript script = Parse("A1:B2 border border bottom thick blue border top left dashed");
        Assert.That(script.Items[0].Elements, Is.EqualTo(new FormatElement[]
        {
            new BorderAction(BorderSides.All, BorderStyle.Thin, new Rgb(0, 0, 0)),
            new BorderAction(BorderSides.Bottom, BorderStyle.Thick, new Rgb(0, 0, 255)),
            new BorderAction(BorderSides.Top | BorderSides.Left, BorderStyle.Dashed, new Rgb(0, 0, 0)),
        }));
    }

    [Test]
    public void FiltersAndConditions()
    {
        FormatScript script = Parse("A1:A9 | nonblank | between -5 2.5 | top 10 % bold cond > $B$1 { fill red stop } cond formula \"MOD(ROW(),2)=0\" { bold }");
        FormatItem item = script.Items[0];
        Assert.That(item.Filters, Is.EqualTo(new Condition[]
        {
            new CellStateCondition(CellState.NonBlank),
            new BetweenCondition(false, new NumberOperand(-5), new NumberOperand(2.5)),
            new RankCondition(true, 10, true),
        }));

        ConditionalRule rule = (ConditionalRule)item.Elements[1];
        Assert.That(rule.Condition, Is.EqualTo(new CompareCondition(CompareOp.Greater, new CellOperand("$B$1"))));
        Assert.That(rule.Stop, Is.True);

        ConditionalRule formula = (ConditionalRule)item.Elements[2];
        Assert.That(formula.Condition, Is.EqualTo(new FormulaCondition("=MOD(ROW(),2)=0")));
    }

    [Test]
    public void CondRulesMixWithStaticActions()
    {
        FormatScript script = Parse("A1:A9 cond < 0 { color red } border all cond scale red yellow green cond databar blue cond icons arrows italic");
        Assert.That(script.Items[0].Elements.Select(e => e.GetType()), Is.EqualTo(new[]
        {
            typeof(ConditionalRule), typeof(BorderAction), typeof(ColorScaleRule), typeof(DataBarRule), typeof(IconSetRule), typeof(ToggleAction),
        }));
    }

    [Test]
    public void CommentsAreSkipped()
    {
        FormatScript script = Parse("/* block\ncomment */ A1 bold // line comment\nA2 italic");
        Assert.That(script.Items, Has.Count.EqualTo(2));
    }

    [Test]
    public void WidthAndHeightNeedMatchingShapes()
    {
        Assert.That(Errors("A1 width 10").Single(), Does.StartWith("1:4: 'width' applies to whole columns"));
        Assert.That(Errors("B:B height 10").Single(), Does.Contain("'height' applies to whole rows"));
        Parse("B:D w 10 A1:* col \"X\" w 5 2:3 h 20 A1:* header h 20 A1:* lastrow h 20");
    }

    [Test]
    public void ConditionalFormatRestrictions()
    {
        Assert.That(Errors("A1 cond > 1 { fs 12 }").Single(), Does.Contain("'fontsize' cannot be used in a conditional format"));
        Assert.That(Errors("A1 cond > 1 { border inside }").Single(), Does.Contain("'inside' is not allowed"));
        Assert.That(Errors("A1 cond > 1 { border thick }").Single(), Does.Contain("not 'thick'"));
        Parse("A1 cond > 1 { border all dotted red bold italic underline strike fmt \"0\" fill blue color white }");
    }

    [Test]
    public void ValueChecks()
    {
        Assert.That(Errors("A1 fill rgb 256 0 0").Single(), Does.Contain("between 0 and 255"));
        Assert.That(Errors("XFE1 bold").Single(), Does.Contain("beyond the last Excel column"));
        Assert.That(Errors("A1048577 bold").Single(), Does.Contain("outside the Excel row range"));
        Assert.That(Errors("A1 | top 101 % bold").Single(), Does.Contain("between 1 and 100%"));
        Assert.That(Errors("A:A w -1").Single(), Does.Contain("cannot be negative"));
    }

    [Test]
    public void SyntaxErrorsHavePositions()
    {
        Assert.That(Errors("A1 bold\nA2 bolder").First(), Does.StartWith("2:4:"));
        Assert.That(Errors("A1 fill").First(), Does.Contain("expecting"));
    }

    [Test]
    public void StringEscapes()
    {
        FormatScript script = Parse("A1 fmt \"say \\\"hi\\\" \\\\ \"");
        Assert.That(script.Items[0].Elements.Single(), Is.EqualTo(new NumberFormatAction("say \"hi\" \\ ")));
        Assert.That(VbaFormatWriter.Str("a\"b\nc"), Is.EqualTo("\"a\"\"b\" & vbLf & \"c\""));
    }

    [Test]
    public void VbaOutput()
    {
        string vba = VbaFormatWriter.Write(Parse(
            "'Data'!A1:* header bold border bottom thin\n" +
            "B2:B9 | > 0 cond between 1 C1 { fill #FF0000 stop }\n" +
            "A:A w 12.5"));

        Assert.That(vba, Does.Contain("Public Sub XlWriteFormat()"));
        Assert.That(vba, Does.Contain("Set ws = ActiveWorkbook.Worksheets(\"Data\")"));
        Assert.That(vba, Does.Contain("XlwTablePart(XlwTable(ws.Range(\"A1\")), \"header\")"));
        Assert.That(vba, Does.Contain("XlwBorder rng, xlEdgeBottom, xlContinuous, xlThin, RGB(0, 0, 0)"));
        Assert.That(vba, Does.Contain("Set rng = Application.Intersect(rng, ws.UsedRange)"));
        Assert.That(vba, Does.Contain("If XlwCompare(v, \">\", 0) Then XlwKeep kept, rs, re, c"));
        Assert.That(vba, Does.Contain("Operator:=xlBetween, Formula1:=\"=1\", Formula2:=XlwCfFormula(\"=C1\", tl)"));
        Assert.That(vba, Does.Contain("fc.StopIfTrue = True"));
        Assert.That(vba, Does.Contain("fc.SetLastPriority"));
        Assert.That(vba, Does.Contain("rng.EntireColumn.ColumnWidth = 12.5"));
    }
}
