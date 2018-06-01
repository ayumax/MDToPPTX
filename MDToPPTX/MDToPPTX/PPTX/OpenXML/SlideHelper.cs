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
    internal class SlideHelper : SlideHelperBase
    {
        private PPTXSlide SlideContent;
        private Dictionary<string, string> ImageIDMap = new Dictionary<string, string>();
        private Dictionary<string, string> HyperLinkIDMap = new Dictionary<string, string>();

        public SlideHelper(PPTXSlide SlideContent)
        {
            this.SlideContent = SlideContent;
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
            if (SlideContent.Title != null)
            {
                AddContent(shapeTree, objectID++, SlideContent.Title, PlaceholderValues.Title);
            }

            var slidePartChildIndex = CreateHyperLinkMap(SlideContent, slidePart1, HyperLinkIDMap);
            CreateImageMap(slidePart1, slidePartChildIndex);

            uint bodyIndex = 1;
            foreach(var bodyContent in SlideContent.TextAreas)
            {
                if (bodyContent.Texts.Count > 0)
                {
                    AddTextBox(shapeTree, objectID++, bodyContent, PlaceholderValues.Body, bodyIndex++);
                }
            }


            foreach (var imageContent in SlideContent.Images)
            {
                if (ImageIDMap.ContainsKey(imageContent.ImageFilePath))
                {
                    AddImageContent(shapeTree, objectID++, imageContent, PlaceholderValues.Picture);
                }
            }

            foreach (var tableContent in SlideContent.Tables)
            {
                AddTableContent(shapeTree, objectID++, tableContent);
            }


            slide1.Save(slidePart1);

            // スライドレイアウトの設定
            var slideMaster = presentationPart.SlideMasterParts.First();
            var slideLayout = slideMaster.GetPartById(SlideContent.SlideLayout.ID);
            slidePart1.AddPart(slideLayout);

            SetSlideID(presentationPart, slidePart1);

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

        private void AddContent(ShapeTree shapeTree1, uint ObjectID, PPTXTextArea Content, PlaceholderValues PlaceHolderType, uint PlaceHolderIndex = uint.MaxValue)
        {
            Shape shape1 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties() { Id = ObjectID, Name = $"Content{ObjectID}" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties() { TextBox = true};
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties1.Append(shapeLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape1 = new PlaceholderShape();// { Type = PlaceHolderType };
            if (PlaceHolderIndex != uint.MaxValue)
            {
                placeholderShape1.Index = PlaceHolderIndex;
            }

            applicationNonVisualDrawingProperties2.Append(placeholderShape1);

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties2);
            ShapeProperties shapeProperties1 = new ShapeProperties();

            A.SolidFill solidFill1 = new A.SolidFill();

            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 20000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 80000 };

            schemeColor1.Append(luminanceModulation1);
            schemeColor1.Append(luminanceOffset1);

            solidFill1.Append(schemeColor1);

            shapeProperties1.Append(solidFill1);

            TextBody textBody1 = new TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);

            A.Transform2D transform2D25 = CreateTransform2D(Content.Transform);
            if (transform2D25 != null)
            {
                shapeProperties1.Append(transform2D25);
            }

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(textBody1);

            foreach (var _textLine in Content.Texts)
            {
                var paragraph = new A.Paragraph(CrateParagraphProperties(_textLine));

                foreach(var _textRun in _textLine.Texts)
                {
                    paragraph.Append(new A.Run()
                    {
                        RunProperties = CreateRunProperties(_textRun, HyperLinkIDMap),
                        Text = new A.Text(_textRun.Text)
                    });
                }
                
                shape1.TextBody.Append(paragraph);
            }

            shapeTree1.Append(shape1);
        }


        private void AddTextBox(ShapeTree shapeTree1, uint ObjectID, PPTXTextArea Content, PlaceholderValues PlaceHolderType, uint PlaceHolderIndex = uint.MaxValue)
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

            A.Transform2D transform2D25 = CreateTransform2D(Content.Transform);
            if (transform2D25 != null)
            {
                shapeProperties1.Append(transform2D25);
            }

            shapeProperties1.Append(presetGeometry1);

            if (Content.BackgroundColor.IsTransparent == false)
            {
                A.SolidFill solidFill1 = new A.SolidFill();
                solidFill1.Append(CreateRGBColorModeHex(Content.BackgroundColor));
                shapeProperties1.Append(solidFill1);
            }


            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(textBody1);

            foreach (var _textLine in Content.Texts)
            {
                var paragraph = new A.Paragraph(CrateParagraphProperties(_textLine));

                foreach (var _textRun in _textLine.Texts)
                {
                    paragraph.Append(new A.Run()
                    {
                        RunProperties = CreateRunProperties(_textRun, HyperLinkIDMap),
                        Text = new A.Text(_textRun.Text)
                    });
                }

                shape1.TextBody.Append(paragraph);
            }

            shapeTree1.Append(shape1);
        }

        private void AddImageContent(ShapeTree shapeTree1, uint ObjectID, PPTXImage Content, PlaceholderValues PlaceHolderType, uint PlaceHolderIndex = uint.MaxValue)
        {
            Picture picture3 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties3 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties83 = new NonVisualDrawingProperties() { Id = ObjectID, Name = $"Content{ObjectID}" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList5 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension5 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement5 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7646098D-2B56-48B2-B526-D64C2464A2F4}\" />");

            nonVisualDrawingPropertiesExtension5.Append(openXmlUnknownElement5);

            nonVisualDrawingPropertiesExtensionList5.Append(nonVisualDrawingPropertiesExtension5);

            nonVisualDrawingProperties83.Append(nonVisualDrawingPropertiesExtensionList5);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties3 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks3 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties3.Append(pictureLocks3);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties83 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties3.Append(nonVisualDrawingProperties83);
            nonVisualPictureProperties3.Append(nonVisualPictureDrawingProperties3);
            nonVisualPictureProperties3.Append(applicationNonVisualDrawingProperties83);

            BlipFill blipFill3 = new BlipFill();

            A.Blip blip3 = new A.Blip() { Embed = ImageIDMap[Content.ImageFilePath] };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip3.Append(blipExtensionList1);

            A.Stretch stretch3 = new A.Stretch();
            A.FillRectangle fillRectangle3 = new A.FillRectangle();

            stretch3.Append(fillRectangle3);

            blipFill3.Append(blip3);
            blipFill3.Append(stretch3);

            ShapeProperties shapeProperties70 = new ShapeProperties();

            
            A.PresetGeometry presetGeometry10 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList10 = new A.AdjustValueList();

            presetGeometry10.Append(adjustValueList10);

            A.Transform2D transform2D25 = CreateTransform2D(Content.Transform);
            if (transform2D25 != null)
            {
                shapeProperties70.Append(transform2D25);
            }

            shapeProperties70.Append(presetGeometry10);

            picture3.Append(nonVisualPictureProperties3);
            picture3.Append(blipFill3);
            picture3.Append(shapeProperties70);

            shapeTree1.Append(picture3);
        }

        private void AddTableContent(ShapeTree shapeTree1, uint ObjectID, PPTXTable Content)
        {
            TableSlideHelper helper = new TableSlideHelper();
            helper.AddContent(shapeTree1, ObjectID, Content, HyperLinkIDMap);
        }

        private void CreateImageMap(SlidePart slidePart1, int slidePartChildLastIndex)
        {
            for (int i = 0; i < SlideContent.Images.Count; ++i)
            {
                var imageFilePath = SlideContent.Images[i].ImageFilePath;
                if (System.IO.File.Exists(imageFilePath) == false)
                {
                    continue;
                }

                if (ImageIDMap.ContainsKey(imageFilePath))
                {
                    continue;
                }

                var fileExt = Path.GetExtension(imageFilePath).ToLower();
                var mime = "text/plain";
                switch (fileExt)
                {
                    case ".png":
                        mime = "image/png";
                        break;
                    case ".jpeg":
                    case ".jpg":
                        mime = "image/jpeg";
                        break;
                    case ".bmp":
                        mime = "image/bmp";
                        break;
                    case ".gif":
                        mime = "image/gif";
                        break;
                }

                var imageID = $"rId{i + slidePartChildLastIndex}";

                ImagePart imagePart1 = slidePart1.AddNewPart<ImagePart>(mime, imageID);
                using (System.IO.FileStream stream = new System.IO.FileStream(imageFilePath, System.IO.FileMode.Open))
                {
                    imagePart1.FeedData(stream);
                }

                ImageIDMap.Add(imageFilePath, imageID);
            }
        }

        
    }

}
