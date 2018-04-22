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
        public static SlidePart CreateSlidePart(PresentationPart presentationPart, string SlideID, string Title)
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
                                new P.ShapeProperties()
                                {
                                    Transform2D = new Transform2D()
                                    {
                                        Offset = new Offset() { X = 0, Y = 0 },
                                        Extents = new Extents()
                                        {
                                            Cx = 9144000,
                                            Cy = 6858000,
                                        }
                                    },
                                },
                                new P.TextBody(
                                    new BodyProperties()
                                    {
                                        Anchor = D.TextAnchoringTypeValues.Center
                                    },
                                    new ListStyle(),
                                    new Paragraph(
                                        new ParagraphProperties()
                                        {
                                            Alignment = D.TextAlignmentTypeValues.Center
                                        }, 
                                        new Run()     
                                        {
                                            Text = new D.Text(Title),
                                        },
                                        new EndParagraphRunProperties() { Language = "ja-JP" })
                                        )
                                    )
                                    )
                                    ),
                    new ColorMapOverride(new MasterColorMapping()));

            //A.BodyProperties bodyProperties3 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Center };
            //A.ListStyle listStyle3 = new A.ListStyle();

            //A.Paragraph paragraph3 = new A.Paragraph();
            //A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            return slidePart1;
        }
    }
}
