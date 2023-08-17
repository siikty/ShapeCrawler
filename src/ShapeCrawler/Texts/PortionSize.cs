﻿using System.Collections.Generic;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;
using ShapeCrawler.Services.Factories;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal record PortionSize : IFontSize
{
    private readonly A.Text aText;
    private readonly IParagraphPortion parentParagraphPortion;

    internal PortionSize(A.Text aText, IParagraphPortion parentParagraphPortion)
    {
        this.aText = aText;
        this.parentParagraphPortion = parentParagraphPortion;
    }
    
    public int Size()
    {
        var fontSize = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.FontSize
            ?.Value;
        if (fontSize != null)
        {
            return fontSize.Value / 100;
        }

        IShape ancestorShape = this.parentParagraphPortion.AncestorShape();
        int ancestorParaLevel = this.parentParagraphPortion.AncestorParagraphLevel();

        if (ancestorShape is { Placeholder: not null } shape)
        {
            if (TryFromPlaceholder(shape, ancestorParaLevel, out var sizeFromPlaceholder))
            {
                return sizeFromPlaceholder;
            }
        }

        if (this.paraLvlToFontData.TryGetValue(ancestorParaLevel, out var fontData))
        {
            if (fontData.FontSize is not null)
            {
                return fontData.FontSize / 100;
            }
        }

        return SCConstants.DefaultFontSize;
    }

    public void Update(int points)
    {
        var parent = this.aText.Parent!;
        var aRunPr = parent.GetFirstChild<A.RunProperties>();
        if (aRunPr == null)
        {
            var builder = new ARunPropertiesBuilder();
            aRunPr = builder.Build();
            parent.InsertAt(aRunPr, 0);
        }

        aRunPr.FontSize = points * 100;
    }

    private static bool TryFromPlaceholder(IShape scShape, int paraLevel, out int i)
    {
        i = -1;
        var placeholder = scShape.Placeholder as SCSlidePlaceholder;
        var referencedShape = placeholder?.ReferencedShape.Value as SlideAutoShape;
        var fontDataPlaceholder = new FontData();
        if (referencedShape != null)
        {
            referencedShape.FillFontData(paraLevel, ref fontDataPlaceholder);
            if (fontDataPlaceholder.FontSize is not null)
            {
                i = fontDataPlaceholder.FontSize / 100;
                return true;
            }
        }

        var slideMaster = scShape.SlideMasterInternal;
        if (placeholder?.Type == SCPlaceholderType.Title)
        {
            var pTextStyles = slideMaster.PSlideMaster.TextStyles!;
            var titleFontSize = pTextStyles.TitleStyle!.Level1ParagraphProperties!
                .GetFirstChild<A.DefaultRunProperties>() !.FontSize!.Value;
            i = titleFontSize / 100;
            return true;
        }

        if (slideMaster.TryGetFontSizeFromBody(paraLevel, out var fontSizeBody))
        {
            i = fontSizeBody / 100;
            return true;
        }

        if (slideMaster.TryGetFontSizeFromOther(paraLevel, out var fontSizeOther))
        {
            i = fontSizeOther / 100;
            return true;
        }

        return false;
    }
}