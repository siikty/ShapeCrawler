﻿using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a collection of paragraphs.
/// </summary>
public interface IParagraphs : IReadOnlyList<IParagraph>
{
    /// <summary>
    ///     Adds a new paragraph in collection.
    /// </summary>
    void Add();
}

internal readonly struct Paragraphs : IParagraphs
{
    private readonly OpenXmlPart sdkOpenXmlPart;
    private readonly OpenXmlElement sdkTextBody;

    internal Paragraphs(OpenXmlPart sdkOpenXmlPart, OpenXmlElement sdkTextBody)
    {
        this.sdkOpenXmlPart = sdkOpenXmlPart;
        this.sdkTextBody = sdkTextBody;
    }

    #region Public Properties

    public int Count => this.ParagraphsCore().Count;
    
    public IParagraph this[int index] => this.ParagraphsCore()[index];
    
    public IEnumerator<IParagraph> GetEnumerator() => this.ParagraphsCore().GetEnumerator();
    
    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    #endregion Public Properties

    public void Add()
    {
        var lastAParagraph = this.sdkTextBody.Elements<A.Paragraph>().Last();
        var newAParagraph = (A.Paragraph)lastAParagraph.CloneNode(true);
        newAParagraph.ParagraphProperties ??= new A.ParagraphProperties();
        lastAParagraph.InsertAfterSelf(newAParagraph);
    }

    private List<Paragraph> ParagraphsCore()
    {
        var aParagraphs = this.sdkTextBody.Elements<A.Paragraph>().ToList();
        if (!aParagraphs.Any())
        {
            return new List<Paragraph>(0);
        }

        var paraList = new List<Paragraph>();
        foreach (var aPara in aParagraphs)
        {
            var para = new Paragraph(this.sdkOpenXmlPart, aPara);
            paraList.Add(para);
        }

        return paraList;
    }
}