using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Excel;
using ShapeCrawler.Exceptions;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts;

internal sealed class Chart : Shape, IChart
{
    private readonly ChartType chartType;
    private readonly Lazy<OpenXmlElement?> firstSeries;
    private readonly P.GraphicFrame pGraphicFrame;
    private readonly ChartPart sdkChartPart;
    private readonly C.PlotArea cPlotArea;

    // Contains chart elements, e.g. <c:pieChart>, <c:barChart>, <c:lineChart> etc. If the chart type is not a combination,
    // then collection contains only single item.
    private readonly IEnumerable<OpenXmlElement> cXCharts;

    private string? chartTitle;

    internal Chart(
        OpenXmlPart sdkOpenXmlPart, 
        ChartPart sdkChartPart, 
        P.GraphicFrame pGraphicFrame,
        IReadOnlyList<ICategory> categories)
        : base(sdkOpenXmlPart,pGraphicFrame)
    {
        this.sdkChartPart = sdkChartPart;
        this.pGraphicFrame = pGraphicFrame;
        this.Categories = categories;
        this.firstSeries = new Lazy<OpenXmlElement?>(this.GetFirstSeries);
        this.cPlotArea = sdkChartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
        this.cXCharts = this.cPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        var pShapeProperties = sdkChartPart.ChartSpace.GetFirstChild<C.ShapeProperties>() !;
        this.Outline = new SlideShapeOutline(sdkOpenXmlPart, pShapeProperties);
        this.Fill = new ShapeFill(sdkOpenXmlPart, pShapeProperties);
        this.SeriesList = new SeriesList(
            sdkChartPart,
            this.cPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal)));
    }

    public ChartType Type
    {
        get
        {
            if (this.cXCharts.Count() > 1)
            {
                return ChartType.Combination;
            }

            var chartName = this.cXCharts.Single().LocalName;
            Enum.TryParse(chartName, true, out ChartType enumChartType);

            return enumChartType;
        }
    }

    public override ShapeType ShapeType => ShapeType.Chart;
    
    public override IShapeOutline Outline { get; }
    
    public override IShapeFill Fill { get; }

    public bool HasTitle
    {
        get
        {
            this.chartTitle ??= this.GetTitleOrDefault();
            return this.chartTitle != null;
        }
    }

    public string? Title
    {
        get
        {
            this.chartTitle = this.GetTitleOrDefault();
            return this.chartTitle;
        }
    }

    public bool HasCategories => false;

    public IReadOnlyList<ICategory> Categories { get; }
    
    public ISeriesList SeriesList { get; }
    
    public bool HasXValues => this.ParseXValues() != null;

    public List<double> XValues
    {
        get
        {
            if (this.ParseXValues() == null)
            {
                throw new NotSupportedException(ExceptionMessages.NotXValues);
            }

            return this.ParseXValues() !;
        }
    }
    
    public override Geometry GeometryType => Geometry.Rectangle;
    
    public IAxesManager Axes => this.GetAxes();
    
    public override bool Removeable => true;
    
    public byte[] BookByteArray() => new ExcelBook(this.sdkChartPart).AsByteArray();
    
    public override void Remove() => this.pGraphicFrame.Remove();
    
    private IAxesManager GetAxes() => new AxesManager(this.cPlotArea);

    private string? GetTitleOrDefault()
    {
        var cTitle = this.sdkChartPart.ChartSpace.GetFirstChild<C.Chart>() !.Title;
        if (cTitle == null)
        {
            // chart has not title
            return null;
        }

        var cChartText = cTitle.ChartText;
        bool staticAvailable = this.TryGetStaticTitle(cChartText!, out var staticTitle);
        if (staticAvailable)
        {
            return staticTitle;
        }

        // Dynamic title
        if (cChartText != null)
        {
            return cChartText.Descendants<C.StringPoint>().Single().InnerText;
        }

        // PieChart uses only one series for view.
        // However, it can have store multiple series data in the spreadsheet.
        if (this.Type == ChartType.PieChart)
        {
            return ((SeriesList)this.SeriesList).First().Name;
        }

        return null;
    }

    private bool TryGetStaticTitle(C.ChartText chartText, out string? staticTitle)
    {
        staticTitle = null;
        if (this.Type == ChartType.Combination)
        {
            staticTitle = chartText.RichText!.Descendants<A.Text>().Select(t => t.Text)
                .Aggregate((t1, t2) => t1 + t2);
            return true;
        }

        var rRich = chartText?.RichText;
        if (rRich != null)
        {
            staticTitle = rRich.Descendants<A.Text>().Select(t => t.Text).Aggregate((t1, t2) => t1 + t2);
            return true;
        }

        return false;
    }

    private List<double>? ParseXValues()
    {
        var cXValues = this.firstSeries.Value?.GetFirstChild<C.XValues>();
        if (cXValues?.NumberReference == null)
        {
            return null;
        }

        if (cXValues.NumberReference.NumberingCache != null)
        {
            var cNumericValues = cXValues.NumberReference.NumberingCache.Descendants<C.NumericValue>();
            var cachedPointValues = new List<double>(cNumericValues.Count());
            foreach (var numericValue in cNumericValues)
            {
                var number = double.Parse(numericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                var roundNumber = Math.Round(number, 1);
                cachedPointValues.Add(roundNumber);
            }

            return cachedPointValues;
        }

        return new ExcelBook(this.sdkChartPart).FormulaValues(cXValues.NumberReference.Formula!.Text);
    }

    private OpenXmlElement? GetFirstSeries()
    {
        return this.cXCharts.First().ChildElements
            .FirstOrDefault(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
    }

    public ChartPart GetChartPart()
    {
        return this.sdkChartPart;
    }
}