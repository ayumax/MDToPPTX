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

            var slidePartChildIndex = CreateHyperLinkMap(slidePart1);
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
            GraphicFrame graphicFrame1 = new GraphicFrame();

            NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new NonVisualGraphicFrameProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties() { Id = ObjectID, Name = "表 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7AB8EDC7-F9EF-4752-9A46-413B9437344B}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoGrouping = true };

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = new ApplicationNonVisualDrawingPropertiesExtensionList();

            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = new ApplicationNonVisualDrawingPropertiesExtension() { Uri = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}" };

            P14.ModificationId modificationId1 = new P14.ModificationId() { Val = (UInt32Value)833561296U };
            modificationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            applicationNonVisualDrawingPropertiesExtension1.Append(modificationId1);

            applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

            applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties2);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);
            nonVisualGraphicFrameProperties1.Append(applicationNonVisualDrawingProperties2);

            Transform transform1 = new Transform();
            A.Offset offset2 = new A.Offset() { X = 2032000L, Y = 719666L };
            A.Extents extents2 = new A.Extents() { Cx = 8127999L, Cy = 1483360L };

            transform1.Append(offset2);
            transform1.Append(extents2);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };


            A.Table table1 = new A.Table();

            A.TableProperties tableProperties1 = new A.TableProperties() { FirstRow = true, BandRow = true };
            A.TableStyleId tableStyleId1 = new A.TableStyleId();
            tableStyleId1.Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

            tableProperties1.Append(tableStyleId1);

            A.TableGrid tableGrid1 = new A.TableGrid();

            A.GridColumn gridColumn1 = new A.GridColumn() { Width = 2709333L };

            A.ExtensionList extensionList1 = new A.ExtensionList();

            A.Extension extension1 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3243622648\" />");

            extension1.Append(openXmlUnknownElement2);

            extensionList1.Append(extension1);

            gridColumn1.Append(extensionList1);

            A.GridColumn gridColumn2 = new A.GridColumn() { Width = 2709333L };

            A.ExtensionList extensionList2 = new A.ExtensionList();

            A.Extension extension2 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"371540315\" />");

            extension2.Append(openXmlUnknownElement3);

            extensionList2.Append(extension2);

            gridColumn2.Append(extensionList2);

            A.GridColumn gridColumn3 = new A.GridColumn() { Width = 2709333L };

            A.ExtensionList extensionList3 = new A.ExtensionList();

            A.Extension extension3 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            OpenXmlUnknownElement openXmlUnknownElement4 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"1972170560\" />");

            extension3.Append(openXmlUnknownElement4);

            extensionList3.Append(extension3);

            gridColumn3.Append(extensionList3);

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);

            A.TableRow tableRow1 = new A.TableRow() { Height = 370840L };

            A.TableCell tableCell1 = new A.TableCell();

            A.TextBody textBody1 = new A.TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();

            A.Run run1 = new A.Run();
            A.RunProperties runProperties1 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text1 = new A.Text();
            text1.Text = "1";

            run1.Append(runProperties1);
            run1.Append(text1);

            A.Run run2 = new A.Run();
            A.RunProperties runProperties2 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };
            A.Text text2 = new A.Text();
            text2.Text = "列目";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph1.Append(run1);
            paragraph1.Append(run2);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);
            A.TableCellProperties tableCellProperties1 = new A.TableCellProperties();

            tableCell1.Append(textBody1);
            tableCell1.Append(tableCellProperties1);

            A.TableCell tableCell2 = new A.TableCell();

            A.TextBody textBody2 = new A.TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties();
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();

            A.Run run3 = new A.Run();
            A.RunProperties runProperties3 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text3 = new A.Text();
            text3.Text = "2";

            run3.Append(runProperties3);
            run3.Append(text3);

            A.Run run4 = new A.Run();
            A.RunProperties runProperties4 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };
            A.Text text4 = new A.Text();
            text4.Text = "列目";

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph2.Append(run3);
            paragraph2.Append(run4);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);
            A.TableCellProperties tableCellProperties2 = new A.TableCellProperties();

            tableCell2.Append(textBody2);
            tableCell2.Append(tableCellProperties2);

            A.TableCell tableCell3 = new A.TableCell();

            A.TextBody textBody3 = new A.TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties();
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();

            A.Run run5 = new A.Run();
            A.RunProperties runProperties5 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text5 = new A.Text();
            text5.Text = "3";

            run5.Append(runProperties5);
            run5.Append(text5);

            A.Run run6 = new A.Run();
            A.RunProperties runProperties6 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };
            A.Text text6 = new A.Text();
            text6.Text = "列目";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph3.Append(run5);
            paragraph3.Append(run6);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph3);
            A.TableCellProperties tableCellProperties3 = new A.TableCellProperties();

            tableCell3.Append(textBody3);
            tableCell3.Append(tableCellProperties3);

            A.ExtensionList extensionList4 = new A.ExtensionList();

            A.Extension extension4 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            OpenXmlUnknownElement openXmlUnknownElement5 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2927081048\" />");

            extension4.Append(openXmlUnknownElement5);

            extensionList4.Append(extension4);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(extensionList4);

            A.TableRow tableRow2 = new A.TableRow() { Height = 370840L };

            A.TableCell tableCell4 = new A.TableCell();

            A.TextBody textBody4 = new A.TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties();
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();

            A.Run run7 = new A.Run();
            A.RunProperties runProperties7 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text7 = new A.Text();
            text7.Text = "1-1";

            run7.Append(runProperties7);
            run7.Append(text7);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };

            paragraph4.Append(run7);
            paragraph4.Append(endParagraphRunProperties1);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph4);
            A.TableCellProperties tableCellProperties4 = new A.TableCellProperties();

            tableCell4.Append(textBody4);
            tableCell4.Append(tableCellProperties4);

            A.TableCell tableCell5 = new A.TableCell();

            A.TextBody textBody5 = new A.TextBody();
            A.BodyProperties bodyProperties5 = new A.BodyProperties();
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.Run run8 = new A.Run();
            A.RunProperties runProperties8 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text8 = new A.Text();
            text8.Text = "1-2";

            run8.Append(runProperties8);
            run8.Append(text8);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };

            paragraph5.Append(run8);
            paragraph5.Append(endParagraphRunProperties2);

            textBody5.Append(bodyProperties5);
            textBody5.Append(listStyle5);
            textBody5.Append(paragraph5);
            A.TableCellProperties tableCellProperties5 = new A.TableCellProperties();

            tableCell5.Append(textBody5);
            tableCell5.Append(tableCellProperties5);

            A.TableCell tableCell6 = new A.TableCell();

            A.TextBody textBody6 = new A.TextBody();
            A.BodyProperties bodyProperties6 = new A.BodyProperties();
            A.ListStyle listStyle6 = new A.ListStyle();

            A.Paragraph paragraph6 = new A.Paragraph();

            A.Run run9 = new A.Run();
            A.RunProperties runProperties9 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text9 = new A.Text();
            text9.Text = "1-3";

            run9.Append(runProperties9);
            run9.Append(text9);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };

            paragraph6.Append(run9);
            paragraph6.Append(endParagraphRunProperties3);

            textBody6.Append(bodyProperties6);
            textBody6.Append(listStyle6);
            textBody6.Append(paragraph6);
            A.TableCellProperties tableCellProperties6 = new A.TableCellProperties();

            tableCell6.Append(textBody6);
            tableCell6.Append(tableCellProperties6);

            A.ExtensionList extensionList5 = new A.ExtensionList();

            A.Extension extension5 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            OpenXmlUnknownElement openXmlUnknownElement6 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"270367390\" />");

            extension5.Append(openXmlUnknownElement6);

            extensionList5.Append(extension5);

            tableRow2.Append(tableCell4);
            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);
            tableRow2.Append(extensionList5);

            A.TableRow tableRow3 = new A.TableRow() { Height = 370840L };

            A.TableCell tableCell7 = new A.TableCell();

            A.TextBody textBody7 = new A.TextBody();
            A.BodyProperties bodyProperties7 = new A.BodyProperties();
            A.ListStyle listStyle7 = new A.ListStyle();

            A.Paragraph paragraph7 = new A.Paragraph();

            A.Run run10 = new A.Run();
            A.RunProperties runProperties10 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text10 = new A.Text();
            text10.Text = "2-1";

            run10.Append(runProperties10);
            run10.Append(text10);
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };

            paragraph7.Append(run10);
            paragraph7.Append(endParagraphRunProperties4);

            textBody7.Append(bodyProperties7);
            textBody7.Append(listStyle7);
            textBody7.Append(paragraph7);
            A.TableCellProperties tableCellProperties7 = new A.TableCellProperties();

            tableCell7.Append(textBody7);
            tableCell7.Append(tableCellProperties7);

            A.TableCell tableCell8 = new A.TableCell();

            A.TextBody textBody8 = new A.TextBody();
            A.BodyProperties bodyProperties8 = new A.BodyProperties();
            A.ListStyle listStyle8 = new A.ListStyle();

            A.Paragraph paragraph8 = new A.Paragraph();

            A.Run run11 = new A.Run();
            A.RunProperties runProperties11 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text11 = new A.Text();
            text11.Text = "2-2";

            run11.Append(runProperties11);
            run11.Append(text11);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };

            paragraph8.Append(run11);
            paragraph8.Append(endParagraphRunProperties5);

            textBody8.Append(bodyProperties8);
            textBody8.Append(listStyle8);
            textBody8.Append(paragraph8);
            A.TableCellProperties tableCellProperties8 = new A.TableCellProperties();

            tableCell8.Append(textBody8);
            tableCell8.Append(tableCellProperties8);

            A.TableCell tableCell9 = new A.TableCell();

            A.TextBody textBody9 = new A.TextBody();
            A.BodyProperties bodyProperties9 = new A.BodyProperties();
            A.ListStyle listStyle9 = new A.ListStyle();

            A.Paragraph paragraph9 = new A.Paragraph();

            A.Run run12 = new A.Run();
            A.RunProperties runProperties12 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text12 = new A.Text();
            text12.Text = "2-3";

            run12.Append(runProperties12);
            run12.Append(text12);
            A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };

            paragraph9.Append(run12);
            paragraph9.Append(endParagraphRunProperties6);

            textBody9.Append(bodyProperties9);
            textBody9.Append(listStyle9);
            textBody9.Append(paragraph9);
            A.TableCellProperties tableCellProperties9 = new A.TableCellProperties();

            tableCell9.Append(textBody9);
            tableCell9.Append(tableCellProperties9);

            A.ExtensionList extensionList6 = new A.ExtensionList();

            A.Extension extension6 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            OpenXmlUnknownElement openXmlUnknownElement7 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3961662475\" />");

            extension6.Append(openXmlUnknownElement7);

            extensionList6.Append(extension6);

            tableRow3.Append(tableCell7);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);
            tableRow3.Append(extensionList6);

            A.TableRow tableRow4 = new A.TableRow() { Height = 370840L };

            A.TableCell tableCell10 = new A.TableCell();

            A.TextBody textBody10 = new A.TextBody();
            A.BodyProperties bodyProperties10 = new A.BodyProperties();
            A.ListStyle listStyle10 = new A.ListStyle();

            A.Paragraph paragraph10 = new A.Paragraph();

            A.Run run13 = new A.Run();
            A.RunProperties runProperties13 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text13 = new A.Text();
            text13.Text = "3-1";

            run13.Append(runProperties13);
            run13.Append(text13);
            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };

            paragraph10.Append(run13);
            paragraph10.Append(endParagraphRunProperties7);

            textBody10.Append(bodyProperties10);
            textBody10.Append(listStyle10);
            textBody10.Append(paragraph10);
            A.TableCellProperties tableCellProperties10 = new A.TableCellProperties();

            tableCell10.Append(textBody10);
            tableCell10.Append(tableCellProperties10);

            A.TableCell tableCell11 = new A.TableCell();

            A.TextBody textBody11 = new A.TextBody();
            A.BodyProperties bodyProperties11 = new A.BodyProperties();
            A.ListStyle listStyle11 = new A.ListStyle();

            A.Paragraph paragraph11 = new A.Paragraph();

            A.Run run14 = new A.Run();
            A.RunProperties runProperties14 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP", Dirty = false };
            A.Text text14 = new A.Text();
            text14.Text = "3-2";

            run14.Append(runProperties14);
            run14.Append(text14);
            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };

            paragraph11.Append(run14);
            paragraph11.Append(endParagraphRunProperties8);

            textBody11.Append(bodyProperties11);
            textBody11.Append(listStyle11);
            textBody11.Append(paragraph11);
            A.TableCellProperties tableCellProperties11 = new A.TableCellProperties();

            tableCell11.Append(textBody11);
            tableCell11.Append(tableCellProperties11);

            A.TableCell tableCell12 = new A.TableCell();

            A.TextBody textBody12 = new A.TextBody();
            A.BodyProperties bodyProperties12 = new A.BodyProperties();
            A.ListStyle listStyle12 = new A.ListStyle();

            A.Paragraph paragraph12 = new A.Paragraph();

            A.Run run15 = new A.Run();
            A.RunProperties runProperties15 = new A.RunProperties() { Kumimoji = true, Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text15 = new A.Text();
            text15.Text = "3-3";

            run15.Append(runProperties15);
            run15.Append(text15);
            A.EndParagraphRunProperties endParagraphRunProperties9 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", Dirty = false };

            paragraph12.Append(run15);
            paragraph12.Append(endParagraphRunProperties9);

            textBody12.Append(bodyProperties12);
            textBody12.Append(listStyle12);
            textBody12.Append(paragraph12);
            A.TableCellProperties tableCellProperties12 = new A.TableCellProperties();

            tableCell12.Append(textBody12);
            tableCell12.Append(tableCellProperties12);

            A.ExtensionList extensionList7 = new A.ExtensionList();

            A.Extension extension7 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            OpenXmlUnknownElement openXmlUnknownElement8 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2203780241\" />");

            extension7.Append(openXmlUnknownElement8);

            extensionList7.Append(extension7);

            tableRow4.Append(tableCell10);
            tableRow4.Append(tableCell11);
            tableRow4.Append(tableCell12);
            tableRow4.Append(extensionList7);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);

            graphicData1.Append(table1);

            graphic1.Append(graphicData1);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);

            shapeTree1.Append(graphicFrame1);
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

        private int CreateHyperLinkMap(SlidePart slidePart1)
        {
            int lastIndex = 2;

            var textRuns = SlideContent.TextAreas
                .SelectMany(_textArea => _textArea.Texts.SelectMany(_text => _text.Texts))
                .Where(_textRun => _textRun.Link.IsEnable)
                .Select((_textRun, _Index) => new { Link = _textRun.Link, Index = _Index });

            foreach (var linkItem in textRuns)
            {
                if (HyperLinkIDMap.ContainsKey(linkItem.Link.LinkKey))
                {
                    continue;
                }

                var linkId = $"rId{linkItem.Index + 2}";
                lastIndex = linkItem.Index + 2;

                slidePart1.AddHyperlinkRelationship(new System.Uri(linkItem.Link.LinkURL, System.UriKind.Absolute), true, linkId);

                HyperLinkIDMap.Add(linkItem.Link.LinkKey, linkId);
            }

            return lastIndex;
        }

        private A.ParagraphProperties CrateParagraphProperties(PPTXText Content)
        {
            var paragraphPorp = new A.ParagraphProperties();

            //paragraphPorp.LineSpacing = new A.LineSpacing() { SpacingPercent = new A.SpacingPercent() { Val = 110000 } };

            switch (Content.Bullet)
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
            A.RunProperties runProperties3 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", FontSize = (int)(Text.Font.FontSize * 100), Dirty = false };
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = Text.Font.FontFamily, Panose = "020B0604030504040204", PitchFamily = 50, CharacterSet = -128 };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = Text.Font.FontFamily, Panose = "020B0604030504040204", PitchFamily = 50, CharacterSet = -128 };

            if (Text.Font.ForegroundColor.IsTransparent == false)
            {
                A.SolidFill solidFill1 = new A.SolidFill();
                solidFill1.Append(CreateRGBColorModeHex(Text.Font.ForegroundColor));
                runProperties3.Append(solidFill1);
            }

            runProperties3.Append(latinFont1);
            runProperties3.Append(eastAsianFont1);

            if (HyperLinkIDMap.ContainsKey(Text.Link.LinkKey))
            {
                A.HyperlinkOnClick hyperlinkOnClick1 = new A.HyperlinkOnClick() { Id = HyperLinkIDMap[Text.Link.LinkKey] };

                runProperties3.Append(hyperlinkOnClick1);
            }


            return runProperties3;
        }

        private A.RgbColorModelHex CreateRGBColorModeHex(PPTXColor Color) => new A.RgbColorModelHex() { Val = $"{Color.Color.R:X02}{Color.Color.G:X02}{Color.Color.B:X02}" };
    }

}
