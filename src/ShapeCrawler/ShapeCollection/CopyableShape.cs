﻿using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Wrappers;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal abstract class CopyableShape : Shape
{
    internal CopyableShape(OpenXmlPart sdkOpenXmlPart, OpenXmlElement openXmlElement)
        : base(sdkOpenXmlPart, openXmlElement)
    {
    }

    internal virtual void CopyTo(
        int id,
        P.ShapeTree pShapeTree,
        IEnumerable<string> existingShapeNames)
    {
        new PShapeTreeWrap(pShapeTree).Add(this.pShapeTreeElement);
    }
}