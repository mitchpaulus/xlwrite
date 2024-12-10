using System;
using System.Collections.Generic;

namespace xlwrite;

public class FormatListener : XlWriteBaseListener
{
    public readonly List<string> Lines = new();

    public override void EnterItem(XlWriteParser.ItemContext context)
    {
        string? worksheet;
        if (context.selection().STRING() is { } s)
        {
            var text = s.GetText();
            worksheet = text.Substring(1, text.Length - 2);
        }
        else
        {
            worksheet = null;
        }

        foreach (XlWriteParser.ActionContext? action in context.actions().action())
        {
            if (action is XlWriteParser.FillActionExpContext fillAction)
            {
                XlWriteParser.ColorContext? color = fillAction.fillAction().color();
                if (color?.rgbColor() is { } rgbColor)
                {
                    Lines.Add($"{GetVbaSelectionObject(worksheet, context.selection().range())}.Interior.Color = RGB({rgbColor.INT(0)}, {rgbColor.INT(1)}, {rgbColor.INT(2)})");
                }
            }
        }
    }


    private string GetVbaSelectionObject(string? worksheet, XlWriteParser.RangeContext rangeContext)
    {
        return worksheet is null
            ? $"ActiveSheet.Range(\"{rangeContext.GetText()}\")"
            : $"ActiveWorkbook.Worksheets(\"{worksheet}\").Range(\"{rangeContext.GetText()}\")";
    }
}
