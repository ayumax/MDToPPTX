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
    internal class SlideHelper
    {
        private PPTXSlide SlideContent;
        private Dictionary<string, string> ImageIDMap = new Dictionary<string, string>();

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

            uint bodyIndex = 1;
            foreach(var bodyContent in SlideContent.TextAreas)
            {
                if (bodyContent.Texts.Count > 0)
                {
                    AddTextBox(shapeTree, objectID++, bodyContent, PlaceholderValues.Body, bodyIndex++);
                }
            }

            CreateImageMap(slidePart1);

            foreach (var imageContent in SlideContent.Images)
            {
                if (ImageIDMap.ContainsKey(imageContent.ImageFilePath))
                {
                    AddImageContent(shapeTree, objectID++, imageContent, PlaceholderValues.Picture);
                }
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
                        RunProperties = CreateRunProperties(_textRun),
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

            A.SolidFill solidFill1 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FF0000" };

            solidFill1.Append(rgbColorModelHex1);

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
            shapeProperties1.Append(solidFill1);

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
                        RunProperties = CreateRunProperties(_textRun),
                        Text = new A.Text(_textRun.Text)
                    });
                }

                shape1.TextBody.Append(paragraph);
            }

            shapeTree1.Append(shape1);
        }

        private void AddImageContent(ShapeTree shapeTree1, uint ObjectID, PPTXImage Content, PlaceholderValues PlaceHolderType, uint PlaceHolderIndex = uint.MaxValue)
        {
            Picture picture1 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties1 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties78 = new NonVisualDrawingProperties() { Id = ObjectID, Name = $"Content{ObjectID}" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoGrouping = true, NoChangeAspect = false };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties78 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape65 = new PlaceholderShape();// { Type = PlaceHolderType };
            if (PlaceHolderIndex != uint.MaxValue)
            {
                placeholderShape65.Index = PlaceHolderIndex;
            }

            applicationNonVisualDrawingProperties78.Append(placeholderShape65);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties78);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties78);

            BlipFill blipFill1 = new BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = ImageIDMap[Content.ImageFilePath] };

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            ShapeProperties shapeProperties65 = new ShapeProperties();

            A.Transform2D transform2D25 = CreateTransform2D(Content.Transform);
            if (transform2D25 != null)
            {
                shapeProperties65.Append(transform2D25);
            }

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties65);

            shapeTree1.Append(picture1);
        }

        private void SetSlideID(PresentationPart presentationPart, SlidePart slidePart1)
        {
            // Insert the new slide into the slide list after the previous slide.
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

            SlideId newSlideId = slideIdList.AppendChild(new SlideId());
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart1);
        }

        private A.Transform2D CreateTransform2D(PPTXTransform transform)
        {
            A.Transform2D transform2D25 = null;

            if (transform.AutoLayout == false)
            {
                transform2D25 = new A.Transform2D()
                {
                    Offset = new A.Offset()
                    {
                        X = Utils.GetCmToShapeScale(transform.PositionX),
                        Y = Utils.GetCmToShapeScale(transform.PositionY)
                    },
                    Extents = new A.Extents()
                    {
                        Cx = Utils.GetCmToShapeScale(transform.SizeX),
                        Cy = Utils.GetCmToShapeScale(transform.SizeY)
                    }
                };
            }
           

            return transform2D25;
        }

        private void CreateImageMap(SlidePart slidePart1)
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

                var imageID = $"rId{i + 2}";

                ImagePart imagePart1 = slidePart1.AddNewPart<ImagePart>(mime, imageID);
                using (System.IO.FileStream stream = new System.IO.FileStream(imageFilePath, System.IO.FileMode.Open))
                {
                    imagePart1.FeedData(stream);
                }

                ImageIDMap.Add(imageFilePath, imageID);
            }
        }

        private A.ParagraphProperties CrateParagraphProperties(PPTXText Content)
        {
            var paragraphPorp = new A.ParagraphProperties();

            switch(Content.Bullet)
            {
                case PPTXBullet.None:
                    paragraphPorp.Append(new A.NoBullet());
                    break;

                case PPTXBullet.Circle:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "l" });
                    break;

                case PPTXBullet.Rectangle:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "n" });
                    break;

                case PPTXBullet.Diamond:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "u" });
                    break;

                case PPTXBullet.RectangleBorder:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "p" });
                    break;

                case PPTXBullet.Check:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "ü" });
                    break;

                case PPTXBullet.Arrow:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "Ø" });
                    break;

                case PPTXBullet.MiniCircle:
                    break;

                case PPTXBullet.Number:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "+mj-lt" });
                    paragraphPorp.Append(new A.AutoNumberedBullet() { Type = A.TextAutoNumberSchemeValues.ArabicPeriod });
                    break;

                case PPTXBullet.CircleNumber:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "+mj-ea" });
                    paragraphPorp.Append(new A.AutoNumberedBullet() { Type = A.TextAutoNumberSchemeValues.CircleNumberDoubleBytePlain });
                    break;
            }

            return paragraphPorp;
        }

        private A.RunProperties CreateRunProperties(PPTXTextRun Text)
        {
            A.RunProperties runProperties3 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US", FontSize = (int)(Text.Font.FontSize * 100), Dirty = false };
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = Text.Font.FontFamily, Panose = "020B0604030504040204", PitchFamily = 50, CharacterSet = -128 };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = Text.Font.FontFamily, Panose = "020B0604030504040204", PitchFamily = 50, CharacterSet = -128 };

            runProperties3.Append(latinFont1);
            runProperties3.Append(eastAsianFont1);

            return runProperties3;
        }
    }

}
