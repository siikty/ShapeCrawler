﻿using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

internal class OLEObject : ShapeCollection.Shape
{
    private readonly P.GraphicFrame pGraphicFrame;

    internal OLEObject(OpenXmlPart sdkOpenXmlPart, P.GraphicFrame pGraphicFrame)
        : base(sdkOpenXmlPart, pGraphicFrame)
    {
        this.pGraphicFrame = pGraphicFrame;
        this.Outline = new SlideShapeOutline(sdkOpenXmlPart, pGraphicFrame.Descendants<P.ShapeProperties>().First());
        this.Fill = new ShapeFill(sdkOpenXmlPart, pGraphicFrame.Descendants<P.ShapeProperties>().First());
    }

    public override ShapeType ShapeType => ShapeType.OLEObject;

    public override bool HasOutline => true;
    
    public override IShapeOutline Outline { get; }
    
    public override bool HasFill => true;
    
    public override IShapeFill Fill { get; }
    
    public override bool Removeable => true;
    
    public override void Remove() => this.pGraphicFrame.Remove();
}