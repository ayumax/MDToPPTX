using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace MDToPPTX.PPTX
{
    public class PPTXDocument : IDisposable
    {
        private PresentationDocument presentationDoc;

        public PPTXDocument()
        {

        }

        public PPTXDocument(string FilePath)
        {
            Init(FilePath);
        }

        public void Init(string FilePath)
        {
            // Create a presentation at a specified file path. The presentation document type is pptx, by default.
            presentationDoc = PresentationDocument.Create(FilePath, PresentationDocumentType.Presentation);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            PPTX.DefaultParts.DefaultPresentationParts.CreatePresentationParts(presentationPart);       
        }

        public void AddSlide(PPTXSlide Slide)
        {
            InsertNewSlide(presentationDoc, Slide);
        }

        private void InsertNewSlide(PresentationDocument presentationDocument, PPTXSlide Slide)
        {
            if (presentationDocument == null || Slide == null) return;

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            P.NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new P.NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

            // Declare and instantiate the body shape of the new slide.
            P.Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
            drawingObjectId++;

            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            bodyShape.ShapeProperties = new P.ShapeProperties()
            {
                Transform2D = new Transform2D()
                {
                    Offset = new Offset()
                    {
                        X = Utils.GetCmToShapeScale(2),
                        Y = Utils.GetCmToShapeScale(4)
                    },
                    Extents = new Extents()
                    {
                        Cx = Utils.GetCmToShapeScale(10),
                        Cy = Utils.GetCmToShapeScale(5)
                    }
                }
            };
            // Specify the text of the body shape.
            bodyShape.TextBody = new P.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph(new Run() { Text = new D.Text(Slide.Body) }, new EndParagraphRunProperties() { Language = "ja-JP" }));

            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                    prevSlideId = slideId;
                }
            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.AppendChild(new SlideId());
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
        }

        public void Close()
        {
            presentationDoc?.Close();
        }

        public void Dispose()
        {
            Close();
        }
        
    }
}
