using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Wrappers;
using A = DocumentFormat.OpenXml.Drawing;

internal sealed class ATextWrap
{
    private readonly OpenXmlPart sdkOpenXmlPart;
    private readonly A.Text aText;

    internal ATextWrap(OpenXmlPart sdkOpenXmlPart, A.Text aText)
    {
        this.sdkOpenXmlPart = sdkOpenXmlPart;
        this.aText = aText;
    }

    internal string EastAsianName()
    {
        var aEastAsianFont = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.GetFirstChild<A.EastAsianFont>();
        if (aEastAsianFont != null)
        {
            if (aEastAsianFont.Typeface == "+mj-ea")
            {
                var themeFontScheme = new ThemeFontScheme(this.sdkOpenXmlPart);
                return themeFontScheme.MajorEastAsianFont();
            }

            return aEastAsianFont.Typeface!;
        }
        
        return new ThemeFontScheme(this.sdkOpenXmlPart).MinorEastAsianFont();
    }

    internal void UpdateEastAsianName(string eastAsianFont)
    {
        var aEastAsianFont = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.GetFirstChild<A.EastAsianFont>();
        if (aEastAsianFont != null)
        {
            aEastAsianFont.Typeface = eastAsianFont;
            return;
        }
        
        new ThemeFontScheme(this.sdkOpenXmlPart).UpdateMinorEastAsianFont(eastAsianFont);
    }
}