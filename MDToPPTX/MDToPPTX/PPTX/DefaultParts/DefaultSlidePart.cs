using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace MDToPPTX.PPTX.DefaultParts
{
    class DefaultSlidePart
    {
        public static SlidePart CreateSlidePart(PresentationPart presentationPart, string SlideID)
        {
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>(SlideID);
            slidePart1.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()),
                            new P.Shape(
                                new P.NonVisualShapeProperties(
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                new P.ShapeProperties() { Transform2D = new Transform2D() { Offset = new Offset() { X = 987136L, Y = 1267691L }, Extents = new Extents() { Cx = 4270664L, Cy = 369332L } }, },
                                new P.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph(new Run() { Text = new D.Text("かみのそのあゆまでう!!!") }, new EndParagraphRunProperties() { Language = "ja-JP" }))))),
                    new ColorMapOverride(new MasterColorMapping()));


            return slidePart1;
        }
    }
}
