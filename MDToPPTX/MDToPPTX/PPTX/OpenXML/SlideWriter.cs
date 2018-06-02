using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;

namespace MDToPPTX.PPTX.OpenXML
{
    internal class SlideWriter
    {
        private PPTXSlide SlideContent;
        private PPTXSlideLayoutGroup SlideLayouts;
        private Dictionary<string, string> HyperLinkIDMap = new Dictionary<string, string>();
        private ImageSlideWriter ImageWriter = new ImageSlideWriter();

        public SlideWriter(PPTXSlide SlideContent, PPTXSlideLayoutGroup SlideLayouts)
        {
            this.SlideContent = SlideContent;
            this.SlideLayouts = SlideLayouts;
        }

        public void InsertNewSlide(PresentationDocument presentationDocument)
        {
            var presentationPart = presentationDocument.PresentationPart;
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>();

            Slide slide1 = new Slide();
            slide1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var shapeTree = InitCommonProperty(slide1);

            uint objectID = 4;
            var slidePartChildIndex = SlideWriterHelper.CreateHyperLinkMap(SlideContent, slidePart1, HyperLinkIDMap);

            ImageWriter.CreateImageMap(SlideContent, slidePart1, slidePartChildIndex);

            foreach(var bodyContent in SlideContent.TextAreas)
            {
                if (bodyContent.Texts.Count > 0)
                {
                    AddTextBox(shapeTree, objectID++, bodyContent);
                }
            }

            objectID = ImageWriter.AddImageContents(shapeTree, objectID);

            foreach (var tableContent in SlideContent.Tables)
            {
                AddTableContent(shapeTree, objectID++, tableContent);
            }


            slide1.Save(slidePart1);

            // スライドレイアウトの設定
            var slideMaster = presentationPart.SlideMasterParts.First();
            var slideLayout = slideMaster.GetPartById(SlideLayouts.SlideLayouts[SlideContent.SlideLayout].ID);
            slidePart1.AddPart(slideLayout);

            SlideWriterHelper.SetSlideID(presentationPart, slidePart1);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
        }

        private ShapeTree InitCommonProperty(Slide slide1)
        {
            CommonSlideData commonSlideData1 = new CommonSlideData();

            ShapeTree shapeTree1 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties1 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties1 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties1 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(nonVisualGroupShapeDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(applicationNonVisualDrawingProperties1);

            GroupShapeProperties groupShapeProperties1 = new GroupShapeProperties();

            A.TransformGroup transformGroup1 = new A.TransformGroup();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset1 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup1.Append(offset1);
            transformGroup1.Append(extents1);
            transformGroup1.Append(childOffset1);
            transformGroup1.Append(childExtents1);

            groupShapeProperties1.Append(transformGroup1);

            shapeTree1.Append(nonVisualGroupShapeProperties1);
            shapeTree1.Append(groupShapeProperties1);

            CommonSlideDataExtensionList commonSlideDataExtensionList1 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension1 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId1 = new P14.CreationId() { Val = (UInt32Value)4221661300U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension1.Append(creationId1);

            commonSlideDataExtensionList1.Append(commonSlideDataExtension1);

            commonSlideData1.Append(shapeTree1);
            commonSlideData1.Append(commonSlideDataExtensionList1);

            ColorMapOverride colorMapOverride1 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping1 = new A.MasterColorMapping();

            colorMapOverride1.Append(masterColorMapping1);

            slide1.Append(commonSlideData1);
            slide1.Append(colorMapOverride1);

            return shapeTree1;

        }

        //private void AddContent(ShapeTree shapeTree1, uint ObjectID, PPTXTextArea Content, PlaceholderValues PlaceHolderType, uint PlaceHolderIndex = uint.MaxValue)
        //{
        //    Shape shape1 = new Shape();

        //    NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();

        //    NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties() { Id = ObjectID, Name = $"Content{ObjectID}" };

        //    NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties() { TextBox = true};
        //    A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoGrouping = true };

        //    nonVisualShapeDrawingProperties1.Append(shapeLocks1);

        //    ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();
        //    PlaceholderShape placeholderShape1 = new PlaceholderShape();// { Type = PlaceHolderType };
        //    if (PlaceHolderIndex != uint.MaxValue)
        //    {
        //        placeholderShape1.Index = PlaceHolderIndex;
        //    }

        //    applicationNonVisualDrawingProperties2.Append(placeholderShape1);

        //    nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
        //    nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
        //    nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties2);
        //    ShapeProperties shapeProperties1 = new ShapeProperties();

        //    A.SolidFill solidFill1 = new A.SolidFill();

        //    A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
        //    A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 20000 };
        //    A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 80000 };

        //    schemeColor1.Append(luminanceModulation1);
        //    schemeColor1.Append(luminanceOffset1);

        //    solidFill1.Append(schemeColor1);

        //    shapeProperties1.Append(solidFill1);

        //    TextBody textBody1 = new TextBody();
        //    A.BodyProperties bodyProperties1 = new A.BodyProperties();
        //    A.ListStyle listStyle1 = new A.ListStyle();

        //    textBody1.Append(bodyProperties1);
        //    textBody1.Append(listStyle1);

        //    A.Transform2D transform2D25 = SlideWriterHelper.CreateTransform2D(Content.Transform);
        //    if (transform2D25 != null)
        //    {
        //        shapeProperties1.Append(transform2D25);
        //    }

        //    shape1.Append(nonVisualShapeProperties1);
        //    shape1.Append(shapeProperties1);
        //    shape1.Append(textBody1);

        //    foreach (var _textLine in Content.Texts)
        //    {
        //        var paragraph = new A.Paragraph(SlideWriterHelper.CrateParagraphProperties(_textLine));

        //        foreach(var _textRun in _textLine.Texts)
        //        {
        //            paragraph.Append(new A.Run()
        //            {
        //                RunProperties = SlideWriterHelper.CreateRunProperties(_textRun, HyperLinkIDMap),
        //                Text = new A.Text(_textRun.Text)
        //            });
        //        }
                
        //        shape1.TextBody.Append(paragraph);
        //    }

        //    shapeTree1.Append(shape1);
        //}


        private void AddTextBox(ShapeTree shapeTree1, uint ObjectID, PPTXTextArea Content)
        {
            Shape shape1 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties() { Id = ObjectID, Name = $"Content{ObjectID}" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{5FE2CA47-E73A-450F-9AE0-DF438874E2FB}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties2);

            ShapeProperties shapeProperties1 = new ShapeProperties();


            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            TextBody textBody1 = new TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);

            A.Transform2D transform2D25 = SlideWriterHelper.CreateTransform2D(Content.Transform);
            if (transform2D25 != null)
            {
                shapeProperties1.Append(transform2D25);
            }

            shapeProperties1.Append(presetGeometry1);

            if (Content.BackgroundColor.IsTransparent == false)
            {
                A.SolidFill solidFill1 = new A.SolidFill();
                solidFill1.Append(SlideWriterHelper.CreateRGBColorModeHex(Content.BackgroundColor));
                shapeProperties1.Append(solidFill1);
            }


            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(textBody1);

            foreach (var _textLine in Content.Texts)
            {
                var paragraph = new A.Paragraph(SlideWriterHelper.CrateParagraphProperties(_textLine));

                foreach (var _textRun in _textLine.Texts)
                {
                    paragraph.Append(new A.Run()
                    {
                        RunProperties = SlideWriterHelper.CreateRunProperties(_textRun, HyperLinkIDMap),
                        Text = new A.Text(_textRun.Text)
                    });
                }

                shape1.TextBody.Append(paragraph);
            }

            shapeTree1.Append(shape1);
        }

        

        private void AddTableContent(ShapeTree shapeTree1, uint ObjectID, PPTXTable Content)
        {
            TableSlideWriter helper = new TableSlideWriter();
            helper.AddContent(shapeTree1, ObjectID, Content, HyperLinkIDMap);
        }

        
        
    }

}
