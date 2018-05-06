using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;


namespace MDToPPTX.PPTX.DefaultParts.SlideLayouts
{
    public class SlideLayout_TwoContents : SlideLayoutPartBase
    {
        public override SlideLayoutPart CreateSlideLayoutPart(OpenXmlPartContainer containerPart)
        {
            SlideLayoutPart slideLayoutPart = containerPart.AddNewPart<SlideLayoutPart>(LayoutSetting.ID);

            SlideLayout slideLayout10 = new SlideLayout() { Type = SlideLayoutValues.TwoObjects, Preserve = true };
            slideLayout10.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout10.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout10.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData12 = new CommonSlideData() { Name = LayoutSetting.Name };

            ShapeTree shapeTree12 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties12 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties65 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties12 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties65 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties12.Append(nonVisualDrawingProperties65);
            nonVisualGroupShapeProperties12.Append(nonVisualGroupShapeDrawingProperties12);
            nonVisualGroupShapeProperties12.Append(applicationNonVisualDrawingProperties65);

            GroupShapeProperties groupShapeProperties12 = new GroupShapeProperties();

            A.TransformGroup transformGroup12 = new A.TransformGroup();
            A.Offset offset31 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents31 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset12 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents12 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup12.Append(offset31);
            transformGroup12.Append(extents31);
            transformGroup12.Append(childOffset12);
            transformGroup12.Append(childExtents12);

            groupShapeProperties12.Append(transformGroup12);

            Shape shape54 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties54 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties66 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties54 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks54 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties54.Append(shapeLocks54);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties66 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape54 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties66.Append(placeholderShape54);

            nonVisualShapeProperties54.Append(nonVisualDrawingProperties66);
            nonVisualShapeProperties54.Append(nonVisualShapeDrawingProperties54);
            nonVisualShapeProperties54.Append(applicationNonVisualDrawingProperties66);
            ShapeProperties shapeProperties54 = new ShapeProperties();

            TextBody textBody54 = new TextBody();
            A.BodyProperties bodyProperties54 = new A.BodyProperties();
            A.ListStyle listStyle54 = new A.ListStyle();

            A.Paragraph paragraph82 = new A.Paragraph();

            A.Run run108 = new A.Run();
            A.RunProperties runProperties128 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text128 = new A.Text();
            text128.Text = "マスター タイトルの書式設定";

            run108.Append(runProperties128);
            run108.Append(text128);
            A.EndParagraphRunProperties endParagraphRunProperties48 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph82.Append(run108);
            paragraph82.Append(endParagraphRunProperties48);

            textBody54.Append(bodyProperties54);
            textBody54.Append(listStyle54);
            textBody54.Append(paragraph82);

            shape54.Append(nonVisualShapeProperties54);
            shape54.Append(shapeProperties54);
            shape54.Append(textBody54);

            Shape shape55 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties55 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties67 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties55 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks55 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties55.Append(shapeLocks55);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties67 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape55 = new PlaceholderShape() { Size = PlaceholderSizeValues.Half, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties67.Append(placeholderShape55);

            nonVisualShapeProperties55.Append(nonVisualDrawingProperties67);
            nonVisualShapeProperties55.Append(nonVisualShapeDrawingProperties55);
            nonVisualShapeProperties55.Append(applicationNonVisualDrawingProperties67);

            ShapeProperties shapeProperties55 = new ShapeProperties();

            A.Transform2D transform2D20 = new A.Transform2D();
            A.Offset offset32 = new A.Offset() { X = 628650L, Y = 1825625L };
            A.Extents extents32 = new A.Extents() { Cx = 3886200L, Cy = 4351338L };

            transform2D20.Append(offset32);
            transform2D20.Append(extents32);

            shapeProperties55.Append(transform2D20);

            TextBody textBody55 = new TextBody();
            A.BodyProperties bodyProperties55 = new A.BodyProperties();
            A.ListStyle listStyle55 = new A.ListStyle();

            A.Paragraph paragraph83 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties40 = new A.ParagraphProperties() { Level = 0 };

            A.Run run109 = new A.Run();
            A.RunProperties runProperties129 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text129 = new A.Text();
            text129.Text = "マスター テキストの書式設定";

            run109.Append(runProperties129);
            run109.Append(text129);

            paragraph83.Append(paragraphProperties40);
            paragraph83.Append(run109);

            A.Paragraph paragraph84 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties41 = new A.ParagraphProperties() { Level = 1 };

            A.Run run110 = new A.Run();
            A.RunProperties runProperties130 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text130 = new A.Text();
            text130.Text = "第 ";

            run110.Append(runProperties130);
            run110.Append(text130);

            A.Run run111 = new A.Run();
            A.RunProperties runProperties131 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text131 = new A.Text();
            text131.Text = "2 ";

            run111.Append(runProperties131);
            run111.Append(text131);

            A.Run run112 = new A.Run();
            A.RunProperties runProperties132 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text132 = new A.Text();
            text132.Text = "レベル";

            run112.Append(runProperties132);
            run112.Append(text132);

            paragraph84.Append(paragraphProperties41);
            paragraph84.Append(run110);
            paragraph84.Append(run111);
            paragraph84.Append(run112);

            A.Paragraph paragraph85 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties42 = new A.ParagraphProperties() { Level = 2 };

            A.Run run113 = new A.Run();
            A.RunProperties runProperties133 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text133 = new A.Text();
            text133.Text = "第 ";

            run113.Append(runProperties133);
            run113.Append(text133);

            A.Run run114 = new A.Run();
            A.RunProperties runProperties134 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text134 = new A.Text();
            text134.Text = "3 ";

            run114.Append(runProperties134);
            run114.Append(text134);

            A.Run run115 = new A.Run();
            A.RunProperties runProperties135 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text135 = new A.Text();
            text135.Text = "レベル";

            run115.Append(runProperties135);
            run115.Append(text135);

            paragraph85.Append(paragraphProperties42);
            paragraph85.Append(run113);
            paragraph85.Append(run114);
            paragraph85.Append(run115);

            A.Paragraph paragraph86 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties43 = new A.ParagraphProperties() { Level = 3 };

            A.Run run116 = new A.Run();
            A.RunProperties runProperties136 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text136 = new A.Text();
            text136.Text = "第 ";

            run116.Append(runProperties136);
            run116.Append(text136);

            A.Run run117 = new A.Run();
            A.RunProperties runProperties137 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text137 = new A.Text();
            text137.Text = "4 ";

            run117.Append(runProperties137);
            run117.Append(text137);

            A.Run run118 = new A.Run();
            A.RunProperties runProperties138 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text138 = new A.Text();
            text138.Text = "レベル";

            run118.Append(runProperties138);
            run118.Append(text138);

            paragraph86.Append(paragraphProperties43);
            paragraph86.Append(run116);
            paragraph86.Append(run117);
            paragraph86.Append(run118);

            A.Paragraph paragraph87 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties44 = new A.ParagraphProperties() { Level = 4 };

            A.Run run119 = new A.Run();
            A.RunProperties runProperties139 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text139 = new A.Text();
            text139.Text = "第 ";

            run119.Append(runProperties139);
            run119.Append(text139);

            A.Run run120 = new A.Run();
            A.RunProperties runProperties140 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text140 = new A.Text();
            text140.Text = "5 ";

            run120.Append(runProperties140);
            run120.Append(text140);

            A.Run run121 = new A.Run();
            A.RunProperties runProperties141 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text141 = new A.Text();
            text141.Text = "レベル";

            run121.Append(runProperties141);
            run121.Append(text141);
            A.EndParagraphRunProperties endParagraphRunProperties49 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph87.Append(paragraphProperties44);
            paragraph87.Append(run119);
            paragraph87.Append(run120);
            paragraph87.Append(run121);
            paragraph87.Append(endParagraphRunProperties49);

            textBody55.Append(bodyProperties55);
            textBody55.Append(listStyle55);
            textBody55.Append(paragraph83);
            textBody55.Append(paragraph84);
            textBody55.Append(paragraph85);
            textBody55.Append(paragraph86);
            textBody55.Append(paragraph87);

            shape55.Append(nonVisualShapeProperties55);
            shape55.Append(shapeProperties55);
            shape55.Append(textBody55);

            Shape shape56 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties56 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties68 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties56 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks56 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties56.Append(shapeLocks56);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties68 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape56 = new PlaceholderShape() { Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties68.Append(placeholderShape56);

            nonVisualShapeProperties56.Append(nonVisualDrawingProperties68);
            nonVisualShapeProperties56.Append(nonVisualShapeDrawingProperties56);
            nonVisualShapeProperties56.Append(applicationNonVisualDrawingProperties68);

            ShapeProperties shapeProperties56 = new ShapeProperties();

            A.Transform2D transform2D21 = new A.Transform2D();
            A.Offset offset33 = new A.Offset() { X = 4629150L, Y = 1825625L };
            A.Extents extents33 = new A.Extents() { Cx = 3886200L, Cy = 4351338L };

            transform2D21.Append(offset33);
            transform2D21.Append(extents33);

            shapeProperties56.Append(transform2D21);

            TextBody textBody56 = new TextBody();
            A.BodyProperties bodyProperties56 = new A.BodyProperties();
            A.ListStyle listStyle56 = new A.ListStyle();

            A.Paragraph paragraph88 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties45 = new A.ParagraphProperties() { Level = 0 };

            A.Run run122 = new A.Run();
            A.RunProperties runProperties142 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text142 = new A.Text();
            text142.Text = "マスター テキストの書式設定";

            run122.Append(runProperties142);
            run122.Append(text142);

            paragraph88.Append(paragraphProperties45);
            paragraph88.Append(run122);

            A.Paragraph paragraph89 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties46 = new A.ParagraphProperties() { Level = 1 };

            A.Run run123 = new A.Run();
            A.RunProperties runProperties143 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text143 = new A.Text();
            text143.Text = "第 ";

            run123.Append(runProperties143);
            run123.Append(text143);

            A.Run run124 = new A.Run();
            A.RunProperties runProperties144 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text144 = new A.Text();
            text144.Text = "2 ";

            run124.Append(runProperties144);
            run124.Append(text144);

            A.Run run125 = new A.Run();
            A.RunProperties runProperties145 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text145 = new A.Text();
            text145.Text = "レベル";

            run125.Append(runProperties145);
            run125.Append(text145);

            paragraph89.Append(paragraphProperties46);
            paragraph89.Append(run123);
            paragraph89.Append(run124);
            paragraph89.Append(run125);

            A.Paragraph paragraph90 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties47 = new A.ParagraphProperties() { Level = 2 };

            A.Run run126 = new A.Run();
            A.RunProperties runProperties146 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text146 = new A.Text();
            text146.Text = "第 ";

            run126.Append(runProperties146);
            run126.Append(text146);

            A.Run run127 = new A.Run();
            A.RunProperties runProperties147 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text147 = new A.Text();
            text147.Text = "3 ";

            run127.Append(runProperties147);
            run127.Append(text147);

            A.Run run128 = new A.Run();
            A.RunProperties runProperties148 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text148 = new A.Text();
            text148.Text = "レベル";

            run128.Append(runProperties148);
            run128.Append(text148);

            paragraph90.Append(paragraphProperties47);
            paragraph90.Append(run126);
            paragraph90.Append(run127);
            paragraph90.Append(run128);

            A.Paragraph paragraph91 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties48 = new A.ParagraphProperties() { Level = 3 };

            A.Run run129 = new A.Run();
            A.RunProperties runProperties149 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text149 = new A.Text();
            text149.Text = "第 ";

            run129.Append(runProperties149);
            run129.Append(text149);

            A.Run run130 = new A.Run();
            A.RunProperties runProperties150 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text150 = new A.Text();
            text150.Text = "4 ";

            run130.Append(runProperties150);
            run130.Append(text150);

            A.Run run131 = new A.Run();
            A.RunProperties runProperties151 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text151 = new A.Text();
            text151.Text = "レベル";

            run131.Append(runProperties151);
            run131.Append(text151);

            paragraph91.Append(paragraphProperties48);
            paragraph91.Append(run129);
            paragraph91.Append(run130);
            paragraph91.Append(run131);

            A.Paragraph paragraph92 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties49 = new A.ParagraphProperties() { Level = 4 };

            A.Run run132 = new A.Run();
            A.RunProperties runProperties152 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text152 = new A.Text();
            text152.Text = "第 ";

            run132.Append(runProperties152);
            run132.Append(text152);

            A.Run run133 = new A.Run();
            A.RunProperties runProperties153 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text153 = new A.Text();
            text153.Text = "5 ";

            run133.Append(runProperties153);
            run133.Append(text153);

            A.Run run134 = new A.Run();
            A.RunProperties runProperties154 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text154 = new A.Text();
            text154.Text = "レベル";

            run134.Append(runProperties154);
            run134.Append(text154);
            A.EndParagraphRunProperties endParagraphRunProperties50 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph92.Append(paragraphProperties49);
            paragraph92.Append(run132);
            paragraph92.Append(run133);
            paragraph92.Append(run134);
            paragraph92.Append(endParagraphRunProperties50);

            textBody56.Append(bodyProperties56);
            textBody56.Append(listStyle56);
            textBody56.Append(paragraph88);
            textBody56.Append(paragraph89);
            textBody56.Append(paragraph90);
            textBody56.Append(paragraph91);
            textBody56.Append(paragraph92);

            shape56.Append(nonVisualShapeProperties56);
            shape56.Append(shapeProperties56);
            shape56.Append(textBody56);

            Shape shape57 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties57 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties69 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties57 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks57 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties57.Append(shapeLocks57);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties69 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape57 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties69.Append(placeholderShape57);

            nonVisualShapeProperties57.Append(nonVisualDrawingProperties69);
            nonVisualShapeProperties57.Append(nonVisualShapeDrawingProperties57);
            nonVisualShapeProperties57.Append(applicationNonVisualDrawingProperties69);
            ShapeProperties shapeProperties57 = new ShapeProperties();

            TextBody textBody57 = new TextBody();
            A.BodyProperties bodyProperties57 = new A.BodyProperties();
            A.ListStyle listStyle57 = new A.ListStyle();

            A.Paragraph paragraph93 = new A.Paragraph();

            A.Field field21 = new A.Field() { Id = "{BFFF7C5F-97F4-4B42-9C39-61661A529470}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties155 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };
            runProperties155.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text155 = new A.Text();
            text155.Text = "2018/5/3";

            field21.Append(runProperties155);
            field21.Append(text155);
            A.EndParagraphRunProperties endParagraphRunProperties51 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };

            paragraph93.Append(field21);
            paragraph93.Append(endParagraphRunProperties51);

            textBody57.Append(bodyProperties57);
            textBody57.Append(listStyle57);
            textBody57.Append(paragraph93);

            shape57.Append(nonVisualShapeProperties57);
            shape57.Append(shapeProperties57);
            shape57.Append(textBody57);

            Shape shape58 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties58 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties70 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties58 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks58 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties58.Append(shapeLocks58);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties70 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape58 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties70.Append(placeholderShape58);

            nonVisualShapeProperties58.Append(nonVisualDrawingProperties70);
            nonVisualShapeProperties58.Append(nonVisualShapeDrawingProperties58);
            nonVisualShapeProperties58.Append(applicationNonVisualDrawingProperties70);
            ShapeProperties shapeProperties58 = new ShapeProperties();

            TextBody textBody58 = new TextBody();
            A.BodyProperties bodyProperties58 = new A.BodyProperties();
            A.ListStyle listStyle58 = new A.ListStyle();

            A.Paragraph paragraph94 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties52 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };

            paragraph94.Append(endParagraphRunProperties52);

            textBody58.Append(bodyProperties58);
            textBody58.Append(listStyle58);
            textBody58.Append(paragraph94);

            shape58.Append(nonVisualShapeProperties58);
            shape58.Append(shapeProperties58);
            shape58.Append(textBody58);

            Shape shape59 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties59 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties71 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties59 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks59 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties59.Append(shapeLocks59);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties71 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape59 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties71.Append(placeholderShape59);

            nonVisualShapeProperties59.Append(nonVisualDrawingProperties71);
            nonVisualShapeProperties59.Append(nonVisualShapeDrawingProperties59);
            nonVisualShapeProperties59.Append(applicationNonVisualDrawingProperties71);
            ShapeProperties shapeProperties59 = new ShapeProperties();

            TextBody textBody59 = new TextBody();
            A.BodyProperties bodyProperties59 = new A.BodyProperties();
            A.ListStyle listStyle59 = new A.ListStyle();

            A.Paragraph paragraph95 = new A.Paragraph();

            A.Field field22 = new A.Field() { Id = "{10647DAF-1A54-42E0-9176-57F2D0EA6A3B}", Type = "slidenum" };

            A.RunProperties runProperties156 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };
            runProperties156.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text156 = new A.Text();
            text156.Text = "‹#›";

            field22.Append(runProperties156);
            field22.Append(text156);
            A.EndParagraphRunProperties endParagraphRunProperties53 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };

            paragraph95.Append(field22);
            paragraph95.Append(endParagraphRunProperties53);

            textBody59.Append(bodyProperties59);
            textBody59.Append(listStyle59);
            textBody59.Append(paragraph95);

            shape59.Append(nonVisualShapeProperties59);
            shape59.Append(shapeProperties59);
            shape59.Append(textBody59);

            shapeTree12.Append(nonVisualGroupShapeProperties12);
            shapeTree12.Append(groupShapeProperties12);
            shapeTree12.Append(shape54);
            shapeTree12.Append(shape55);
            shapeTree12.Append(shape56);
            shapeTree12.Append(shape57);
            shapeTree12.Append(shape58);
            shapeTree12.Append(shape59);

            CommonSlideDataExtensionList commonSlideDataExtensionList12 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension12 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId12 = new P14.CreationId() { Val = (UInt32Value)2293587739U };
            creationId12.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension12.Append(creationId12);

            commonSlideDataExtensionList12.Append(commonSlideDataExtension12);

            commonSlideData12.Append(shapeTree12);
            commonSlideData12.Append(commonSlideDataExtensionList12);

            ColorMapOverride colorMapOverride11 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping11 = new A.MasterColorMapping();

            colorMapOverride11.Append(masterColorMapping11);

            slideLayout10.Append(commonSlideData12);
            slideLayout10.Append(colorMapOverride11);

            slideLayoutPart.SlideLayout = slideLayout10;

            return slideLayoutPart;
        }
    }
}
