using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class Picture : CopyableShape, IPicture
{
    private readonly StringValue blipEmbed;
    private readonly P.Picture pPicture;
    private readonly A.Blip aBlip;

    internal Picture(
        OpenXmlPart sdkOpenXmlPart,
        P.Picture pPicture,
        A.Blip aBlip)
        : this(sdkOpenXmlPart, pPicture, aBlip, new SlidePictureImage(sdkOpenXmlPart, aBlip))
    {
    }

    private Picture(OpenXmlPart sdkOpenXmlPart, P.Picture pPicture, A.Blip aBlip, IImage image)
        : base(sdkOpenXmlPart, pPicture)
    {
        this.pPicture = pPicture;
        this.aBlip = aBlip;
        this.Image = image;
        this.blipEmbed = aBlip.Embed!;
        this.Outline = new SlideShapeOutline(sdkOpenXmlPart, pPicture.ShapeProperties!);
        this.Fill = new ShapeFill(sdkOpenXmlPart, pPicture.ShapeProperties!);
    }

    public IImage Image { get; }
   
    public string? SvgContent => this.GetSvgContent();
    
    public override Geometry GeometryType => Geometry.Rectangle;
    
    public override ShapeType ShapeType => ShapeType.Picture;
    
    public override bool HasOutline => true;
    
    public override IShapeOutline Outline { get; }

    public override bool HasFill => true;
    
    public override IShapeFill Fill { get; }
    
    public override bool Removeable => true;
   
    public override void Remove() => this.pPicture.Remove();

    internal override void CopyTo(
        int id, 
        P.ShapeTree pShapeTree, 
        IEnumerable<string> existingShapeNames)
    {
        base.CopyTo(id, pShapeTree, existingShapeNames);

        // COPY PARTS
        var sourceSdkSlidePart = this.sdkOpenXmlPart;
        var sourceImagePart = (ImagePart)sourceSdkSlidePart.GetPartById(this.blipEmbed.Value!);

        // Creates a new part in this slide with a new Id...
        var targetImagePartRId = this.sdkOpenXmlPart.NextRelationshipId();

        // Adds to current slide parts and update relation id.
        var targetImagePart = this.sdkOpenXmlPart.AddNewPart<ImagePart>(sourceImagePart.ContentType, targetImagePartRId);
        using var sourceImageStream = sourceImagePart.GetStream(FileMode.Open);
        sourceImageStream.Position = 0;
        targetImagePart.FeedData(sourceImageStream);

        var copy = this.pShapeTreeElement.CloneNode(true);
        copy.Descendants<A.Blip>().First().Embed = targetImagePartRId;
    }
    
    private string? GetSvgContent()
    {
        var bel = this.aBlip.GetFirstChild<A.BlipExtensionList>();
        var svgBlipList = bel?.Descendants<SVGBlip>();
        if (svgBlipList == null)
        {
            return null;
        }

        var svgId = svgBlipList.First().Embed!.Value!;

        var imagePart = (ImagePart)this.sdkOpenXmlPart.GetPartById(svgId);
        using var svgStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var sReader = new StreamReader(svgStream);

        return sReader.ReadToEnd();
    }
}