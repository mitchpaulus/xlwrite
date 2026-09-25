using System;
using System.Collections.Generic;

namespace xlwrite;

// Intermediate model for the formatting language. The compiler builds it from
// the parse tree and each backend (currently VBA) writes it out.

public record FormatScript(List<FormatItem> Items);

/// <param name="Line">1-based line of the item in the source script.</param>
/// <param name="Elements">Static actions and conditional rules, in source order.</param>
public record FormatItem(int Line, List<Selection> Selections, List<Condition> Filters, List<FormatElement> Elements);

public enum RefKind { Cell, CellRange, Columns, Rows, Table }

public enum TablePart { All, Header, Body, Column, LastRow }

/// <param name="Sheet">Resolved sheet name, or null for the active sheet.</param>
/// <param name="Address">A1 style address. For tables, the anchor cell.</param>
/// <param name="ColumnName">Header text when <paramref name="Part"/> is <see cref="TablePart.Column"/>.</param>
public record Selection(string? Sheet, RefKind Kind, string Address, TablePart Part = TablePart.All, string? ColumnName = null);

public readonly record struct Rgb(int R, int G, int B);

// Operands

public abstract record Operand;
public record NumberOperand(double Value) : Operand;
public record TextOperand(string Value) : Operand;
public record CellOperand(string Address) : Operand;

// Conditions

public enum CompareOp { Equal, NotEqual, Less, LessEqual, Greater, GreaterEqual }
public enum TextOp { Contains, NotContains, BeginsWith, EndsWith }
public enum CellState { Blank, NonBlank, Error, NoError }
public enum DatePeriod { Today, Yesterday, Tomorrow, Last7Days, ThisWeek, LastWeek, NextWeek, ThisMonth, LastMonth, NextMonth }

public abstract record Condition;
public record CompareCondition(CompareOp Op, Operand Value) : Condition;
public record BetweenCondition(bool Negate, Operand Low, Operand High) : Condition;
public record TextCondition(TextOp Op, string Text) : Condition;
public record CellStateCondition(CellState State) : Condition;
public record RankCondition(bool Top, int Rank, bool Percent) : Condition;
public record AverageCondition(bool Above) : Condition;
public record DuplicateCondition(bool Duplicate) : Condition;
public record DatePeriodCondition(DatePeriod Period) : Condition;
public record FormulaCondition(string Formula) : Condition;

// Elements

public abstract record FormatElement;

public abstract record FormatAction : FormatElement;

public enum ToggleProperty { Bold, Italic, Underline, Strike, Wrap }
public enum HAlign { Left, Center, Right }
public enum VAlign { Top, Middle, Bottom }
public enum BorderStyle { Thin, Medium, Thick, Dashed, Dotted, Double, None }

[Flags]
public enum BorderSides
{
    None = 0,
    Top = 1,
    Bottom = 2,
    Left = 4,
    Right = 8,
    InsideHorizontal = 16,
    InsideVertical = 32,
    Outline = Top | Bottom | Left | Right,
    Inside = InsideHorizontal | InsideVertical,
    All = Outline | Inside,
}

public record ToggleAction(ToggleProperty Property, bool On) : FormatAction;
public record FillAction(Rgb Color) : FormatAction;
public record FontColorAction(Rgb Color) : FormatAction;
public record FontNameAction(string Name) : FormatAction;
public record FontSizeAction(double Size) : FormatAction;
public record WidthAction(double Width) : FormatAction;
public record HeightAction(double Height) : FormatAction;
public record NumberFormatAction(string Format) : FormatAction;
public record HAlignAction(HAlign Align) : FormatAction;
public record VAlignAction(VAlign Align) : FormatAction;
public record BorderAction(BorderSides Sides, BorderStyle Style, Rgb Color) : FormatAction;

public abstract record CfRule : FormatElement;
public record ConditionalRule(Condition Condition, List<FormatAction> Actions, bool Stop) : CfRule;
public record ColorScaleRule(List<Rgb> Colors) : CfRule;
public record DataBarRule(Rgb Color) : CfRule;

public enum IconSet { Arrows, Flags, Traffic, Stars, Symbols, Ratings }
public record IconSetRule(IconSet Set) : CfRule;
