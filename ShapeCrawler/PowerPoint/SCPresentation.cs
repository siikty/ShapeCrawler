﻿using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Collections;
using ShapeCrawler.Constants;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     <inheritdoc cref="IPresentation"/>
/// </summary>
[SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — ShapeCrawler")]
public sealed class SCPresentation : IPresentation
{
    private bool closed;
    private readonly Lazy<Dictionary<int, FontData>> paraLvlToFontData;
    private readonly Lazy<SCSlideSize> slideSize;
    private readonly ResettableLazy<SCSectionCollection> sectionCollectionLazy;
    private readonly ResettableLazy<SCSlideCollection> slideCollectionLazy;
    private Stream? outerStream;
    private string? outerPath;
    private readonly MemoryStream internalStream;

    private SCPresentation(string outerPath)
    {
        this.outerPath = outerPath;

        ThrowIfSourceInvalid(outerPath);
        var pptxBytes = File.ReadAllBytes(outerPath);

        this.internalStream = pptxBytes.ToExpandableStream();
        this.SDKPresentationInternal = PresentationDocument.Open(this.internalStream, true);

        this.ThrowIfSlidesNumberLarge();
        this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
        this.SlideMastersValue =
            new ResettableLazy<SlideMasterCollection>(() => SlideMasterCollection.Create(this));
        this.paraLvlToFontData =
            new Lazy<Dictionary<int, FontData>>(() =>
                ParseFontHeights(this.SDKPresentationInternal.PresentationPart!.Presentation));
        this.sectionCollectionLazy =
            new ResettableLazy<SCSectionCollection>(() => SCSectionCollection.Create(this));
        this.slideCollectionLazy = new ResettableLazy<SCSlideCollection>(() => new SCSlideCollection(this));
    }

    private SCPresentation(Stream sourceStream)
    {
        this.outerStream = sourceStream;
        ThrowIfSourceInvalid(sourceStream);

        this.internalStream = new MemoryStream();
        sourceStream.CopyTo(this.internalStream);
        this.SDKPresentationInternal = PresentationDocument.Open(this.internalStream, true);

        this.ThrowIfSlidesNumberLarge();
        this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
        this.SlideMastersValue =
            new ResettableLazy<SlideMasterCollection>(() => SlideMasterCollection.Create(this));
        this.paraLvlToFontData =
            new Lazy<Dictionary<int, FontData>>(() =>
                ParseFontHeights(this.SDKPresentationInternal.PresentationPart!.Presentation));
        this.sectionCollectionLazy =
            new ResettableLazy<SCSectionCollection>(() => SCSectionCollection.Create(this));
        this.slideCollectionLazy = new ResettableLazy<SCSlideCollection>(() => new SCSlideCollection(this));
    }

    private SCPresentation(byte[] sourceBytes)
    {
        this.internalStream = sourceBytes.ToExpandableStream();
        this.SDKPresentationInternal = PresentationDocument.Open(this.internalStream, true);

        this.ThrowIfSlidesNumberLarge();
        this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
        this.SlideMastersValue =
            new ResettableLazy<SlideMasterCollection>(() => SlideMasterCollection.Create(this));
        this.paraLvlToFontData =
            new Lazy<Dictionary<int, FontData>>(() =>
                ParseFontHeights(this.SDKPresentationInternal.PresentationPart!.Presentation));
        this.sectionCollectionLazy =
            new ResettableLazy<SCSectionCollection>(() => SCSectionCollection.Create(this));
        this.slideCollectionLazy = new ResettableLazy<SCSlideCollection>(() => new SCSlideCollection(this));
    }

    /// <inheritdoc/>
    public ISlideCollection Slides => this.slideCollectionLazy.Value;

    /// <inheritdoc/>
    public int SlideWidth => this.slideSize.Value.Width;

    /// <inheritdoc/>
    public int SlideHeight => this.slideSize.Value.Height;

    /// <inheritdoc/>
    public ISlideMasterCollection SlideMasters => this.SlideMastersValue.Value;

    /// <inheritdoc/>
    public byte[] BinaryData => this.GetByteArray();

    /// <inheritdoc/>
    public ISectionCollection Sections => this.sectionCollectionLazy.Value;

    /// <inheritdoc/>
    public PresentationDocument SDKPresentation => this.GetSDKPresentation();

    internal ResettableLazy<SlideMasterCollection> SlideMastersValue { get; private set; }

    internal PresentationDocument SDKPresentationInternal { get; private set; }

    internal SCSectionCollection SectionsInternal => (SCSectionCollection)this.Sections;

    internal List<ChartWorkbook> ChartWorkbooks { get; } = new ();

    internal Dictionary<int, FontData> ParaLvlToFontData => this.paraLvlToFontData.Value;

    internal List<ImagePart> ImageParts => this.GetImageParts();

    internal SCSlideCollection SlidesInternal => (SCSlideCollection)this.Slides;

    #region Public Methods

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    public static IPresentation Create()
    {
        var stream = new MemoryStream();
        var presDoc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation);
        var presPart = presDoc.AddPresentationPart();
        presPart.Presentation = new P.Presentation();

        CreatePresentationParts(presPart);

        presDoc.Close();

        return Open(stream);
    }

    /// <summary>
    ///     Opens existing presentation from specified file path.
    /// </summary>
    public static IPresentation Open(string pptxPath)
    {
        return new SCPresentation(pptxPath);
    }

    /// <summary>
    ///     Opens presentation from specified byte array.
    /// </summary>
    public static IPresentation Open(byte[] pptxBytes)
    {
        ThrowIfSourceInvalid(pptxBytes);

        return new SCPresentation(pptxBytes);
    }

    /// <summary>
    ///     Opens presentation from specified stream.
    /// </summary>
    public static IPresentation Open(Stream pptxStream)
    {
        pptxStream.Position = 0;
        return new SCPresentation(pptxStream);
    }

    /// <inheritdoc/>
    public void Save()
    {
        this.ChartWorkbooks.ForEach(chartWorkbook => chartWorkbook.Close());
        this.SDKPresentationInternal.Save();

        if (this.outerStream != null)
        {
            this.SDKPresentationInternal.Clone(this.outerStream);
        }
        else if (this.outerPath != null)
        {
            var pres = this.SDKPresentationInternal.Clone(this.outerPath);
            pres.Close();
        }
    }

    /// <inheritdoc/>
    public void SaveAs(string path)
    {
        this.outerStream = null;
        this.outerPath = path;
        this.Save();
    }

    /// <inheritdoc/>
    public void SaveAs(Stream stream)
    {
        this.outerPath = null;
        this.outerStream = stream;
        this.Save();
    }

    /// <inheritdoc/>
    public void Close()
    {
        if (this.closed)
        {
            return;
        }

        this.ChartWorkbooks.ForEach(cw => cw.Close());
        this.SDKPresentationInternal.Close();

        this.closed = true;
    }

    /// <summary>
    ///     Closes presentation and releases resources.
    /// </summary>
    public void Dispose()
    {
        this.Close();
    }

    #endregion Public Methods

    private static void CreatePresentationParts(PresentationPart presPart)
    {
        var slideMasterIdList = new P.SlideMasterIdList(new P.SlideMasterId
            { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
        var slideIdList = new P.SlideIdList(new P.SlideId { Id = (UInt32Value)256U, RelationshipId = "rId2" });
        var slideSize = new P.SlideSize { Cx = 9144000, Cy = 6858000, Type = P.SlideSizeValues.Screen4x3 };
        var notesSize = new P.NotesSize { Cx = 6858000, Cy = 9144000 };
        var defaultTextStyle = new P.DefaultTextStyle();

        presPart.Presentation.Append(
            slideMasterIdList,
            slideIdList,
            slideSize,
            notesSize,
            defaultTextStyle);

        var slidePart = presPart.AddNewSlidePart("rId2");
        var slideLayoutPart = CreateSlideLayoutPart(slidePart);
        var slideMasterPart = CreateSlideMasterPart(slideLayoutPart);
        var themePart = CreateTheme(slideMasterPart);

        slideMasterPart.AddPart(slideLayoutPart, "rId1");
        presPart.AddPart(slideMasterPart, "rId1");
        presPart.AddPart(themePart, "rId5");
    }

    private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart)
    {
        var slideLayoutPart = slidePart.AddNewPart<SlideLayoutPart>("rId1");
        var slideLayout = new P.SlideLayout(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = (UInt32Value)1U, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        slideLayoutPart.SlideLayout = slideLayout;

        return slideLayoutPart;
    }

    private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
    {
        var slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
        var slideMaster = new P.SlideMaster(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = (UInt32Value)1U, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup()),
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
                        new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                        new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
                    new P.ShapeProperties(),
                    new P.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph())))),
            new P.ColorMap()
            {
                Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new P.SlideLayoutIdList(new P.SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));
        slideMasterPart1.SlideMaster = slideMaster;

        return slideMasterPart1;
    }

    private static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
    {
        ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
        A.Theme theme1 = new A.Theme() { Name = "Office Theme" };

        A.ThemeElements themeElements1 = new A.ThemeElements(
            new A.ColorScheme(
                new A.Dark1Color(new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                new A.Light1Color(new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                new A.Dark2Color(new A.RgbColorModelHex() { Val = "1F497D" }),
                new A.Light2Color(new A.RgbColorModelHex() { Val = "EEECE1" }),
                new A.Accent1Color(new A.RgbColorModelHex() { Val = "4F81BD" }),
                new A.Accent2Color(new A.RgbColorModelHex() { Val = "C0504D" }),
                new A.Accent3Color(new A.RgbColorModelHex() { Val = "9BBB59" }),
                new A.Accent4Color(new A.RgbColorModelHex() { Val = "8064A2" }),
                new A.Accent5Color(new A.RgbColorModelHex() { Val = "4BACC6" }),
                new A.Accent6Color(new A.RgbColorModelHex() { Val = "F79646" }),
                new A.Hyperlink(new A.RgbColorModelHex() { Val = "0000FF" }),
                new A.FollowedHyperlinkColor(new A.RgbColorModelHex() { Val = "800080" })) { Name = "Office" },
            new A.FontScheme(
                new A.MajorFont(
                    new A.LatinFont() { Typeface = "Calibri" },
                    new A.EastAsianFont() { Typeface = "" },
                    new A.ComplexScriptFont() { Typeface = "" }),
                new A.MinorFont(
                    new A.LatinFont() { Typeface = "Calibri" },
                    new A.EastAsianFont() { Typeface = "" },
                    new A.ComplexScriptFont() { Typeface = "" })) { Name = "Office" },
            new A.FormatScheme(
                new A.FillStyleList(
                    new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
                    new A.GradientFill(
                        new A.GradientStopList(
                            new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 50000 },
                                        new A.SaturationModulation() { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 0 },
                            new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 37000 },
                                        new A.SaturationModulation() { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 35000 },
                            new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 15000 },
                                        new A.SaturationModulation() { Val = 350000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 100000 }),
                        new A.LinearGradientFill() { Angle = 16200000, Scaled = true }),
                    new A.NoFill(),
                    new A.PatternFill(),
                    new A.GroupFill()),
                new A.LineStyleList(
                    new A.Outline(
                        new A.SolidFill(
                            new A.SchemeColor(
                                new A.Shade() { Val = 95000 },
                                new A.SaturationModulation() { Val = 105000 }) { Val = A.SchemeColorValues.PhColor }),
                        new A.PresetDash() { Val = A.PresetLineDashValues.Solid })
                    {
                        Width = 9525,
                        CapType = A.LineCapValues.Flat,
                        CompoundLineType = A.CompoundLineValues.Single,
                        Alignment = A.PenAlignmentValues.Center
                    },
                    new A.Outline(
                        new A.SolidFill(
                            new A.SchemeColor(
                                new A.Shade() { Val = 95000 },
                                new A.SaturationModulation() { Val = 105000 }) { Val = A.SchemeColorValues.PhColor }),
                        new A.PresetDash() { Val = A.PresetLineDashValues.Solid })
                    {
                        Width = 9525,
                        CapType = A.LineCapValues.Flat,
                        CompoundLineType = A.CompoundLineValues.Single,
                        Alignment = A.PenAlignmentValues.Center
                    },
                    new A.Outline(
                        new A.SolidFill(
                            new A.SchemeColor(
                                new A.Shade() { Val = 95000 },
                                new A.SaturationModulation() { Val = 105000 }) { Val = A.SchemeColorValues.PhColor }),
                        new A.PresetDash() { Val = A.PresetLineDashValues.Solid })
                    {
                        Width = 9525,
                        CapType = A.LineCapValues.Flat,
                        CompoundLineType = A.CompoundLineValues.Single,
                        Alignment = A.PenAlignmentValues.Center
                    }),
                new A.EffectStyleList(
                    new A.EffectStyle(
                        new A.EffectList(
                            new A.OuterShadow(
                                new A.RgbColorModelHex(
                                    new A.Alpha() { Val = 38000 }) { Val = "000000" })
                            {
                                BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false
                            })),
                    new A.EffectStyle(
                        new A.EffectList(
                            new A.OuterShadow(
                                new A.RgbColorModelHex(
                                    new A.Alpha() { Val = 38000 }) { Val = "000000" })
                            {
                                BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false
                            })),
                    new A.EffectStyle(
                        new A.EffectList(
                            new A.OuterShadow(
                                new A.RgbColorModelHex(
                                    new A.Alpha() { Val = 38000 }) { Val = "000000" })
                            {
                                BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false
                            }))),
                new A.BackgroundFillStyleList(
                    new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
                    new A.GradientFill(
                        new A.GradientStopList(
                            new A.GradientStop(
                                new A.SchemeColor(new A.Tint() { Val = 50000 },
                                        new A.SaturationModulation() { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                            new A.GradientStop(
                                new A.SchemeColor(new A.Tint() { Val = 50000 },
                                        new A.SaturationModulation() { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                            new A.GradientStop(
                                new A.SchemeColor(new A.Tint() { Val = 50000 },
                                        new A.SaturationModulation() { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor }) { Position = 0 }),
                        new A.LinearGradientFill() { Angle = 16200000, Scaled = true }),
                    new A.GradientFill(
                        new A.GradientStopList(
                            new A.GradientStop(
                                new A.SchemeColor(new A.Tint() { Val = 50000 },
                                        new A.SaturationModulation() { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                            new A.GradientStop(
                                new A.SchemeColor(new A.Tint() { Val = 50000 },
                                        new A.SaturationModulation() { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor }) { Position = 0 }),
                        new A.LinearGradientFill() { Angle = 16200000, Scaled = true }))) { Name = "Office" });

        theme1.Append(themeElements1);
        theme1.Append(new A.ObjectDefaults());
        theme1.Append(new A.ExtraColorSchemeList());

        themePart1.Theme = theme1;
        return themePart1;
    }

    private byte[] GetByteArray()
    {
        var stream = new MemoryStream();
        this.SDKPresentationInternal.Clone(stream);

        return stream.ToArray();
    }

    private PresentationDocument GetSDKPresentation()
    {
        return (PresentationDocument)this.SDKPresentationInternal.Clone();
    }

    private List<ImagePart> GetImageParts()
    {
        var allShapes = this.SlidesInternal.SelectMany(slide => slide.Shapes);
        var imgParts = new List<ImagePart>();

        FromShapes(allShapes);

        return imgParts;

        void FromShapes(IEnumerable<IShape> shapes)
        {
            foreach (var shape in shapes)
            {
                switch (shape)
                {
                    case SlidePicture slidePicture:
                        imgParts.Add(((SCImage)slidePicture.Image).SDKImagePart);
                        break;
                    case IGroupShape groupShape:
                        FromShapes(groupShape.Shapes.Select(x => x));
                        break;
                }
            }
        }
    }

    private static Dictionary<int, FontData> ParseFontHeights(P.Presentation pPresentation)
    {
        var lvlToFontData = new Dictionary<int, FontData>();

        // from presentation default text settings
        if (pPresentation.DefaultTextStyle != null)
        {
            lvlToFontData = FontDataParser.FromCompositeElement(pPresentation.DefaultTextStyle);
        }

        // from theme default text settings
        if (lvlToFontData.Any(kvp => kvp.Value.FontSize is null))
        {
            var themeTextDefault =
                pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault;
            if (themeTextDefault != null)
            {
                lvlToFontData = FontDataParser.FromCompositeElement(themeTextDefault.ListStyle!);
            }
        }

        return lvlToFontData;
    }

    private static void ThrowIfSourceInvalid(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException(nameof(path));
        }

        var fileInfo = new FileInfo(path);

        ThrowIfPptxSizeLarge(fileInfo.Length);
    }

    private static void ThrowIfSourceInvalid(Stream stream)
    {
        ThrowIfPptxSizeLarge(stream.Length);
    }

    private static void ThrowIfSourceInvalid(byte[] bytes)
    {
        ThrowIfPptxSizeLarge(bytes.Length);
    }

    private static void ThrowIfPptxSizeLarge(in long length)
    {
        if (length > Limitations.MaxPresentationSize)
        {
            throw PresentationIsLargeException.FromMax(Limitations.MaxPresentationSize);
        }
    }

    private void ThrowIfSlidesNumberLarge()
    {
        var nbSlides = this.SDKPresentationInternal.PresentationPart!.SlideParts.Count();
        if (nbSlides > Limitations.MaxSlidesNumber)
        {
            this.Close();
            throw SlidesMuchMoreException.FromMax(Limitations.MaxSlidesNumber);
        }
    }

    private SCSlideSize GetSlideSize()
    {
        var pSlideSize = this.SDKPresentationInternal.PresentationPart!.Presentation.SlideSize!;
        var withPx = UnitConverter.HorizontalEmuToPixel(pSlideSize.Cx!.Value);
        var heightPx = UnitConverter.VerticalEmuToPixel(pSlideSize.Cy!.Value);

        return new SCSlideSize(withPx, heightPx);
    }
}