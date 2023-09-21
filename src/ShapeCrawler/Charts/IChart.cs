﻿using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a chart.
/// </summary>
public interface IChart : IShape
{
    /// <summary>
    ///     Gets chart type.
    /// </summary>
    SCChartType Type { get; }

    /// <summary>
    ///     Gets a value indicating whether the chart has a title.
    /// </summary>
    public bool HasTitle { get; }
    
    /// <summary>
    ///     Gets chart title.
    /// </summary>
    string Title { get; }

    /// <summary>
    ///     Gets a value indicating whether the chart has categories.
    /// </summary>
    bool HasCategories { get; }

    /// <summary>
    ///     Gets collection of categories.
    /// </summary>
    public IReadOnlyCollection<ICategory> Categories { get; }

    /// <summary>
    ///     Gets collection of data series.
    /// </summary>
    ISeriesList SeriesList { get; }

    /// <summary>
    ///     Gets a value indicating whether the chart has x-axis values.
    /// </summary>
    bool HasXValues { get; }

    /// <summary>
    ///     Gets collection of x-axis values.
    /// </summary>
    List<double> XValues { get; } // TODO: should be excluded

    /// <summary>
    ///     Gets byte array of workbook containing chart data source.
    /// </summary>
    byte[] WorkbookByteArray { get; }

    /// <summary>
    ///     Gets instance of <see cref="SpreadsheetDocument"/> of Open XML SDK.
    /// </summary>
    SpreadsheetDocument SDKSpreadsheetDocument { get; }

    /// <summary>
    ///     Gets chart axes manager.
    /// </summary>
    IAxesManager Axes { get; }
}