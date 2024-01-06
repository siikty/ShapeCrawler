using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.ShapeCollection;

internal sealed class GroupShape : Shape, IGroupShape
{
    private readonly P.GroupShape pGroupShape;

    internal GroupShape(OpenXmlPart sdkOpenXmlPart, P.GroupShape pGroupShape)
        : base(sdkOpenXmlPart, pGroupShape)
    {
        this.pGroupShape = pGroupShape;
        this.Shapes = new GroupedShapes(sdkOpenXmlPart, pGroupShape.Elements<OpenXmlCompositeElement>());
        this.Outline = new SlideShapeOutline(sdkOpenXmlPart, pGroupShape.Descendants<P.ShapeProperties>().First());
        this.Fill = new ShapeFill(sdkOpenXmlPart, pGroupShape.Descendants<P.ShapeProperties>().First());
    }

    public IShapes Shapes { get; }
    
    public override Geometry GeometryType => Geometry.Rectangle;
    
    public override ShapeType ShapeType => ShapeType.Group;
    
    public override bool HasOutline => true;
    
    public override IShapeOutline Outline { get; }
    
    public override bool HasFill => true;
    
    public override IShapeFill Fill { get; }
    
    public override bool Removeable => true;
    
    public override double Rotation
    {
        get
        {
            var aTransformGroup = this.pGroupShape.GroupShapeProperties!.TransformGroup!;
            var rotation = aTransformGroup.Rotation?.Value ?? 0;
            return rotation / 60000d;
        }
    }
    
    public override void Remove() => this.pGroupShape.Remove();
}