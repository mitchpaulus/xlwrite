using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;

namespace xlwrite;

/// <summary>
/// Builds a <see cref="FormatScript"/> from the parse tree and performs the semantic checks
/// the grammar can't express. Errors are collected as "line:column: message".
/// </summary>
public class FormatCompiler
{
    private const int MaxColumn = 16384;
    private const int MaxRow = 1048576;

    public readonly List<string> Errors = new();

    public static (FormatScript? Script, List<string> Errors) Compile(ICharStream stream)
    {
        XlWriteLexer lexer = new(stream);
        CommonTokenStream tokenStream = new(lexer);
        XlWriteParser parser = new(tokenStream);

        ErrorListener errorListener = new();
        lexer.RemoveErrorListeners();
        lexer.AddErrorListener(errorListener);
        parser.RemoveErrorListeners();
        parser.AddErrorListener(errorListener);

        XlWriteParser.FileContext file = parser.file();
        if (errorListener.Messages.Any()) return (null, errorListener.Messages);

        FormatCompiler compiler = new();
        FormatScript script = compiler.Build(file);
        return compiler.Errors.Any() ? (null, compiler.Errors) : (script, compiler.Errors);
    }

    public static (FormatScript? Script, List<string> Errors) Compile(string source) => Compile(new AntlrInputStream(source));

    public FormatScript Build(XlWriteParser.FileContext file)
    {
        return new FormatScript(file.item().Select(BuildItem).ToList());
    }

    private FormatItem BuildItem(XlWriteParser.ItemContext context)
    {
        List<Selection> selections = new();
        string? currentSheet = null;
        foreach (XlWriteParser.SelectionContext selectionContext in context.selection())
        {
            if (selectionContext.sheet() is { } sheetContext) currentSheet = SheetName(sheetContext);
            selections.Add(BuildSelection(selectionContext.@ref(), currentSheet));
        }

        List<Condition> filters = context.filter().Select(f => BuildCondition(f.condition())).ToList();

        List<FormatElement> elements = new();
        foreach (XlWriteParser.ElementContext elementContext in context.element())
        {
            switch (elementContext)
            {
                case XlWriteParser.ActionElementContext a:
                    FormatAction action = BuildAction(a.action());
                    CheckSelectionShape(a.action(), action, selections);
                    elements.Add(action);
                    break;
                case XlWriteParser.CfElementContext c:
                    elements.Add(BuildCfRule(c.cfRule()));
                    break;
            }
        }

        return new FormatItem(context.Start.Line, selections, filters, elements);
    }

    private string SheetName(XlWriteParser.SheetContext context)
    {
        string name = context switch
        {
            XlWriteParser.QuotedSheetContext q => UnquoteString(q.STRING().GetText()),
            XlWriteParser.PrefixSheetContext p => UnquotePrefix(p.SHEET_PREFIX().GetText()),
            _ => throw new InvalidOperationException(),
        };
        if (name.Length == 0) Error(context, "Sheet name cannot be empty.");
        return name;
    }

    private static string UnquotePrefix(string text)
    {
        // Strip the trailing '!' and, for 'quoted' names, the quotes and doubled apostrophes.
        string name = text[..^1];
        if (name.StartsWith('\'')) name = name[1..^1].Replace("''", "'");
        return name;
    }

    public static string UnquoteString(string text)
    {
        StringBuilder b = new();
        string inner = text[1..^1];
        for (int i = 0; i < inner.Length; i++)
        {
            if (inner[i] == '\\' && i + 1 < inner.Length && inner[i + 1] is '"' or '\\')
            {
                b.Append(inner[i + 1]);
                i++;
            }
            else
            {
                b.Append(inner[i]);
            }
        }
        return b.ToString();
    }

    private Selection BuildSelection(XlWriteParser.RefContext context, string? sheet)
    {
        switch (context)
        {
            case XlWriteParser.CellRefContext c:
                return new Selection(sheet, RefKind.Cell, CheckedAddress(c.CELL()));
            case XlWriteParser.CellRangeRefContext r:
                return new Selection(sheet, RefKind.CellRange, CheckedAddress(r.CELL_RANGE()));
            case XlWriteParser.ColumnRangeRefContext col:
                return new Selection(sheet, RefKind.Columns, CheckedAddress(col.COL_RANGE()));
            case XlWriteParser.RowRangeRefContext row:
                return new Selection(sheet, RefKind.Rows, CheckedAddress(row.ROW_RANGE()));
            case XlWriteParser.TableRefContext t:
            {
                string anchor = CheckedAddress(t.TABLE())[..^2]; // Drop ':*'
                return t.tablePart() switch
                {
                    null => new Selection(sheet, RefKind.Table, anchor),
                    XlWriteParser.TableHeaderContext => new Selection(sheet, RefKind.Table, anchor, TablePart.Header),
                    XlWriteParser.TableBodyContext => new Selection(sheet, RefKind.Table, anchor, TablePart.Body),
                    XlWriteParser.TableLastRowContext => new Selection(sheet, RefKind.Table, anchor, TablePart.LastRow),
                    XlWriteParser.TableColumnContext tc => new Selection(sheet, RefKind.Table, anchor, TablePart.Column, UnquoteString(tc.STRING().GetText())),
                    _ => throw new InvalidOperationException(),
                };
            }
            default:
                throw new InvalidOperationException();
        }
    }

    private static readonly Regex AddressPart = new(@"\$?([A-Za-z]*)\$?([0-9]*)");

    /// <summary>
    /// Uppercases the address and checks each column and row is within Excel's limits.
    /// </summary>
    private string CheckedAddress(ITerminalNode node)
    {
        string text = node.GetText().ToUpperInvariant();
        foreach (string part in text.Split(':'))
        {
            if (part == "*") continue;
            Match m = AddressPart.Match(part);
            string letters = m.Groups[1].Value;
            string digits = m.Groups[2].Value;
            if (letters.Length > 0 && letters.ExcelColumnNameToInt() > MaxColumn)
            {
                Error(node.Symbol, $"Column '{letters}' is beyond the last Excel column XFD.");
            }
            if (digits.Length > 0 && (!int.TryParse(digits, out int row) || row < 1 || row > MaxRow))
            {
                Error(node.Symbol, $"Row '{digits}' is outside the Excel row range 1 to {MaxRow}.");
            }
        }
        return text;
    }

    private FormatAction BuildAction(XlWriteParser.ActionContext context)
    {
        switch (context)
        {
            case XlWriteParser.BoldActionContext a: return new ToggleAction(ToggleProperty.Bold, Toggle(a.toggle()));
            case XlWriteParser.ItalicActionContext a: return new ToggleAction(ToggleProperty.Italic, Toggle(a.toggle()));
            case XlWriteParser.UnderlineActionContext a: return new ToggleAction(ToggleProperty.Underline, Toggle(a.toggle()));
            case XlWriteParser.StrikeActionContext a: return new ToggleAction(ToggleProperty.Strike, Toggle(a.toggle()));
            case XlWriteParser.WrapActionContext a: return new ToggleAction(ToggleProperty.Wrap, Toggle(a.toggle()));
            case XlWriteParser.FillActionContext a: return new FillAction(Color(a.colorValue()));
            case XlWriteParser.FontColorActionContext a: return new FontColorAction(Color(a.colorValue()));
            case XlWriteParser.FontActionContext a: return new FontNameAction(UnquoteString(a.STRING().GetText()));
            case XlWriteParser.FontSizeActionContext a: return new FontSizeAction(PositiveNumber(a.number(), "Font size"));
            case XlWriteParser.WidthActionContext a: return new WidthAction(PositiveNumber(a.number(), "Width"));
            case XlWriteParser.HeightActionContext a: return new HeightAction(PositiveNumber(a.number(), "Height"));
            case XlWriteParser.NumberFormatActionContext a: return new NumberFormatAction(UnquoteString(a.STRING().GetText()));
            case XlWriteParser.AlignActionContext a:
                return new HAlignAction(a.hAlign().GetText() switch
                {
                    "left" => HAlign.Left,
                    "center" => HAlign.Center,
                    _ => HAlign.Right,
                });
            case XlWriteParser.VAlignActionContext a:
                return new VAlignAction(a.vAlign().GetText() switch
                {
                    "top" => VAlign.Top,
                    "middle" or "center" => VAlign.Middle,
                    _ => VAlign.Bottom,
                });
            case XlWriteParser.BorderActionContext a:
            {
                BorderSides sides = BorderSides.None;
                foreach (XlWriteParser.BorderSideContext side in a.borderSide())
                {
                    sides |= side.GetText() switch
                    {
                        "all" => BorderSides.All,
                        "outline" => BorderSides.Outline,
                        "inside" => BorderSides.Inside,
                        "top" => BorderSides.Top,
                        "bottom" => BorderSides.Bottom,
                        "left" => BorderSides.Left,
                        _ => BorderSides.Right,
                    };
                }
                if (sides == BorderSides.None) sides = BorderSides.All;

                BorderStyle style = a.borderStyle()?.GetText() switch
                {
                    null or "thin" => BorderStyle.Thin,
                    "medium" => BorderStyle.Medium,
                    "thick" => BorderStyle.Thick,
                    "dashed" => BorderStyle.Dashed,
                    "dotted" => BorderStyle.Dotted,
                    "double" => BorderStyle.Double,
                    _ => BorderStyle.None,
                };

                Rgb color = a.colorValue() is { } c ? Color(c) : new Rgb(0, 0, 0);
                return new BorderAction(sides, style, color);
            }
            default:
                throw new InvalidOperationException($"Unhandled action '{context.GetText()}'.");
        }
    }

    private void CheckSelectionShape(XlWriteParser.ActionContext context, FormatAction action, List<Selection> selections)
    {
        if (action is WidthAction)
        {
            foreach (Selection s in selections.Where(s => s.Kind != RefKind.Columns && !(s.Kind == RefKind.Table && s.Part == TablePart.Column)))
            {
                Error(context, $"'width' applies to whole columns, but '{Describe(s)}' is not a column selection. Use a column range like B:D or a table column like A1:* col \"Name\".");
            }
        }
        else if (action is HeightAction)
        {
            foreach (Selection s in selections.Where(s => s.Kind != RefKind.Rows && !(s.Kind == RefKind.Table && s.Part is TablePart.Header or TablePart.LastRow)))
            {
                Error(context, $"'height' applies to whole rows, but '{Describe(s)}' is not a row selection. Use a row range like 2:5, or a table's header or lastrow.");
            }
        }
    }

    private static string Describe(Selection s) => s.Kind switch
    {
        RefKind.Table => s.Part switch
        {
            TablePart.All => $"{s.Address}:*",
            TablePart.Column => $"{s.Address}:* col \"{s.ColumnName}\"",
            _ => $"{s.Address}:* {s.Part.ToString().ToLowerInvariant()}",
        },
        _ => s.Address,
    };

    private CfRule BuildCfRule(XlWriteParser.CfRuleContext context)
    {
        switch (context)
        {
            case XlWriteParser.CondRuleContext c:
            {
                List<FormatAction> actions = new();
                foreach (XlWriteParser.ActionContext actionContext in c.action())
                {
                    FormatAction action = BuildAction(actionContext);
                    CheckConditionalAction(actionContext, action);
                    actions.Add(action);
                }
                return new ConditionalRule(BuildCondition(c.condition()), actions, c.STOP() is not null);
            }
            case XlWriteParser.ScaleRuleContext s:
                return new ColorScaleRule(s.colorValue().Select(Color).ToList());
            case XlWriteParser.DataBarRuleContext d:
                return new DataBarRule(Color(d.colorValue()));
            case XlWriteParser.IconsRuleContext i:
                return new IconSetRule(i.iconSet().GetText() switch
                {
                    "arrows" => IconSet.Arrows,
                    "flags" => IconSet.Flags,
                    "traffic" => IconSet.Traffic,
                    "stars" => IconSet.Stars,
                    "symbols" => IconSet.Symbols,
                    _ => IconSet.Ratings,
                });
            default:
                throw new InvalidOperationException();
        }
    }

    /// <summary>
    /// Excel conditional formats can only change font style and color, fill, border and number format.
    /// </summary>
    private void CheckConditionalAction(XlWriteParser.ActionContext context, FormatAction action)
    {
        string? name = action switch
        {
            WidthAction => "width",
            HeightAction => "height",
            FontNameAction => "font",
            FontSizeAction => "fontsize",
            HAlignAction => "align",
            VAlignAction => "valign",
            ToggleAction { Property: ToggleProperty.Wrap } => "wrap",
            _ => null,
        };
        if (name is not null)
        {
            Error(context, $"'{name}' cannot be used in a conditional format. Excel only allows font style and color, fill, border, and number format.");
        }

        if (action is BorderAction border)
        {
            if ((border.Sides & BorderSides.Inside) != 0 && (border.Sides & BorderSides.Outline) != BorderSides.Outline)
            {
                Error(context, "Conditional borders apply to each cell's edges, so 'inside' is not allowed. Use 'all', 'outline', or individual sides.");
            }
            if (border.Style is BorderStyle.Medium or BorderStyle.Thick or BorderStyle.Double)
            {
                Error(context, $"Conditional borders only support thin, dashed, dotted, and none line styles, not '{border.Style.ToString().ToLowerInvariant()}'.");
            }
        }
    }

    private Condition BuildCondition(XlWriteParser.ConditionContext context)
    {
        switch (context)
        {
            case XlWriteParser.CompareCondContext c:
            {
                CompareOp op = c.compareOp().GetText() switch
                {
                    "=" => CompareOp.Equal,
                    "!=" or "<>" => CompareOp.NotEqual,
                    "<" => CompareOp.Less,
                    "<=" => CompareOp.LessEqual,
                    ">" => CompareOp.Greater,
                    _ => CompareOp.GreaterEqual,
                };
                return new CompareCondition(op, BuildOperand(c.operand()));
            }
            case XlWriteParser.BetweenCondContext b:
                return new BetweenCondition(false, BuildOperand(b.operand(0)), BuildOperand(b.operand(1)));
            case XlWriteParser.NotBetweenCondContext b:
                return new BetweenCondition(true, BuildOperand(b.operand(0)), BuildOperand(b.operand(1)));
            case XlWriteParser.ContainsCondContext t: return new TextCondition(TextOp.Contains, UnquoteString(t.STRING().GetText()));
            case XlWriteParser.NotContainsCondContext t: return new TextCondition(TextOp.NotContains, UnquoteString(t.STRING().GetText()));
            case XlWriteParser.BeginsCondContext t: return new TextCondition(TextOp.BeginsWith, UnquoteString(t.STRING().GetText()));
            case XlWriteParser.EndsCondContext t: return new TextCondition(TextOp.EndsWith, UnquoteString(t.STRING().GetText()));
            case XlWriteParser.BlankCondContext: return new CellStateCondition(CellState.Blank);
            case XlWriteParser.NonBlankCondContext: return new CellStateCondition(CellState.NonBlank);
            case XlWriteParser.ErrorCondContext: return new CellStateCondition(CellState.Error);
            case XlWriteParser.NoErrorCondContext: return new CellStateCondition(CellState.NoError);
            case XlWriteParser.TopCondContext t: return BuildRank(t, true, t.INT(), t.PERCENT() is not null);
            case XlWriteParser.BottomCondContext t: return BuildRank(t, false, t.INT(), t.PERCENT() is not null);
            case XlWriteParser.AboveAvgCondContext: return new AverageCondition(true);
            case XlWriteParser.BelowAvgCondContext: return new AverageCondition(false);
            case XlWriteParser.DuplicateCondContext: return new DuplicateCondition(true);
            case XlWriteParser.UniqueCondContext: return new DuplicateCondition(false);
            case XlWriteParser.DateCondContext d:
                return new DatePeriodCondition(d.datePeriod().GetText() switch
                {
                    "today" => DatePeriod.Today,
                    "yesterday" => DatePeriod.Yesterday,
                    "tomorrow" => DatePeriod.Tomorrow,
                    "last7days" => DatePeriod.Last7Days,
                    "thisweek" => DatePeriod.ThisWeek,
                    "lastweek" => DatePeriod.LastWeek,
                    "nextweek" => DatePeriod.NextWeek,
                    "thismonth" => DatePeriod.ThisMonth,
                    "lastmonth" => DatePeriod.LastMonth,
                    _ => DatePeriod.NextMonth,
                });
            case XlWriteParser.FormulaCondContext f:
            {
                string formula = UnquoteString(f.STRING().GetText()).Trim();
                if (formula.Length == 0) Error(f, "Formula cannot be empty.");
                return new FormulaCondition(formula.StartsWith('=') ? formula : "=" + formula);
            }
            default:
                throw new InvalidOperationException($"Unhandled condition '{context.GetText()}'.");
        }
    }

    private RankCondition BuildRank(ParserRuleContext context, bool top, ITerminalNode rankNode, bool percent)
    {
        int max = percent ? 100 : 1000;
        if (!int.TryParse(rankNode.GetText(), out int rank) || rank < 1 || rank > max)
        {
            Error(context, $"Rank must be between 1 and {max}{(percent ? "%" : "")}.");
        }
        return new RankCondition(top, rank, percent);
    }

    private Operand BuildOperand(XlWriteParser.OperandContext context) => context switch
    {
        XlWriteParser.NumberOperandContext n => new NumberOperand(Number(n.number())),
        XlWriteParser.StringOperandContext s => new TextOperand(UnquoteString(s.STRING().GetText())),
        XlWriteParser.CellOperandContext c => new CellOperand(CheckedAddress(c.CELL())),
        _ => throw new InvalidOperationException(),
    };

    private static bool Toggle(XlWriteParser.ToggleContext? context) => context?.GetText() is null or "on" or "true";

    private static double Number(XlWriteParser.NumberContext context)
    {
        return double.Parse(context.GetText(), NumberStyles.Float, CultureInfo.InvariantCulture);
    }

    private double PositiveNumber(XlWriteParser.NumberContext context, string what)
    {
        double value = Number(context);
        if (value < 0) Error(context, $"{what} cannot be negative.");
        return value;
    }

    private Rgb Color(XlWriteParser.ColorValueContext context)
    {
        switch (context)
        {
            case XlWriteParser.HexColorContext h:
            {
                string hex = h.HEXCOLOR().GetText()[1..];
                return new Rgb(Convert.ToInt32(hex[..2], 16), Convert.ToInt32(hex[2..4], 16), Convert.ToInt32(hex[4..], 16));
            }
            case XlWriteParser.RgbColorContext r:
            {
                int[] parts = new int[3];
                for (int i = 0; i < 3; i++)
                {
                    if (!int.TryParse(r.INT(i).GetText(), out parts[i]) || parts[i] > 255)
                    {
                        Error(r.INT(i).Symbol, $"RGB component '{r.INT(i).GetText()}' must be between 0 and 255.");
                    }
                }
                return new Rgb(parts[0], parts[1], parts[2]);
            }
            case XlWriteParser.NamedColorContext n:
                return n.knownColor().GetText() switch
                {
                    "red" => new Rgb(255, 0, 0),
                    "green" => new Rgb(0, 128, 0),
                    "blue" => new Rgb(0, 0, 255),
                    "black" => new Rgb(0, 0, 0),
                    "white" => new Rgb(255, 255, 255),
                    "gray" or "grey" => new Rgb(128, 128, 128),
                    "orange" => new Rgb(255, 165, 0),
                    "yellow" => new Rgb(255, 255, 0),
                    "purple" => new Rgb(128, 0, 128),
                    var other => throw new InvalidOperationException($"Unhandled color '{other}'."),
                };
            default:
                throw new InvalidOperationException();
        }
    }

    private void Error(ParserRuleContext context, string message) => Error(context.Start, message);

    private void Error(IToken token, string message) => Errors.Add($"{token.Line}:{token.Column + 1}: {message}");
}
