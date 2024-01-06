﻿// ReSharper disable once CheckNamespace

using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

using A = DocumentFormat.OpenXml.Drawing;

/// <summary>
///     Represents a theme.
/// </summary>
public interface ITheme
{
    /// <summary>
    ///     Gets font scheme.
    /// </summary>
    IThemeFontScheme FontScheme { get; }

    /// <summary>
    ///     Gets color scheme.
    /// </summary>
    IThemeColorScheme ColorScheme { get; }
}

internal sealed class Theme : ITheme
{
    private readonly OpenXmlPart sdkOpenXmlPart;
    private readonly A.Theme aTheme;

    internal Theme(OpenXmlPart sdkOpenXmlPart, A.Theme aTheme)
    {
        this.sdkOpenXmlPart = sdkOpenXmlPart;
        this.aTheme = aTheme;
    }

    public IThemeFontScheme FontScheme => new ThemeFontScheme(this.sdkOpenXmlPart);

    public IThemeColorScheme ColorScheme => this.GetColorScheme();

    private IThemeColorScheme GetColorScheme()
    {
        return new ThemeColorScheme(this.aTheme.ThemeElements!.ColorScheme!);
    }
}