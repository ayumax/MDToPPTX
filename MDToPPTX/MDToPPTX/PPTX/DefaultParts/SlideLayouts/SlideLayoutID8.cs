using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;


namespace MDToPPTX.PPTX.DefaultParts.SlideLayouts
{
    public class SlideLayoutID8 : SlideLayoutPartBase
    {
        public override SlideLayoutPart CreateSlideLayoutPart(OpenXmlPartContainer containerPart, string ID)
        {
            SlideLayoutPart slideLayoutPart = containerPart.AddNewPart<SlideLayoutPart>(ID);

            SlideLayout slideLayout2 = new SlideLayout() { Type = SlideLayoutValues.ObjectText, Preserve = true };
            slideLayout2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout2.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData4 = new CommonSlideData() { Name = "タイトル付きのコンテンツ" };

            ShapeTree shapeTree4 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties4 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties16 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties4 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties16 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties4.Append(nonVisualDrawingProperties16);
            nonVisualGroupShapeProperties4.Append(nonVisualGroupShapeDrawingProperties4);
            nonVisualGroupShapeProperties4.Append(applicationNonVisualDrawingProperties16);

            GroupShapeProperties groupShapeProperties4 = new GroupShapeProperties();

            A.TransformGroup transformGroup4 = new A.TransformGroup();
            A.Offset offset11 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents11 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset4 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents4 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup4.Append(offset11);
            transformGroup4.Append(extents11);
            transformGroup4.Append(childOffset4);
            transformGroup4.Append(childExtents4);

            groupShapeProperties4.Append(transformGroup4);

            Shape shape13 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties13 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties17 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties13 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks13 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties13.Append(shapeLocks13);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties17 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape13 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties17.Append(placeholderShape13);

            nonVisualShapeProperties13.Append(nonVisualDrawingProperties17);
            nonVisualShapeProperties13.Append(nonVisualShapeDrawingProperties13);
            nonVisualShapeProperties13.Append(applicationNonVisualDrawingProperties17);

            ShapeProperties shapeProperties13 = new ShapeProperties();

            A.Transform2D transform2D8 = new A.Transform2D();
            A.Offset offset12 = new A.Offset() { X = 629841L, Y = 457200L };
            A.Extents extents12 = new A.Extents() { Cx = 2949178L, Cy = 1600200L };

            transform2D8.Append(offset12);
            transform2D8.Append(extents12);

            shapeProperties13.Append(transform2D8);

            TextBody textBody13 = new TextBody();
            A.BodyProperties bodyProperties13 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle13 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties9 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties34 = new A.DefaultRunProperties() { FontSize = 3200 };

            level1ParagraphProperties9.Append(defaultRunProperties34);

            listStyle13.Append(level1ParagraphProperties9);

            A.Paragraph paragraph17 = new A.Paragraph();

            A.Run run19 = new A.Run();
            A.RunProperties runProperties23 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text23 = new A.Text();
            text23.Text = "マスター タイトルの書式設定";

            run19.Append(runProperties23);
            run19.Append(text23);
            A.EndParagraphRunProperties endParagraphRunProperties11 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph17.Append(run19);
            paragraph17.Append(endParagraphRunProperties11);

            textBody13.Append(bodyProperties13);
            textBody13.Append(listStyle13);
            textBody13.Append(paragraph17);

            shape13.Append(nonVisualShapeProperties13);
            shape13.Append(shapeProperties13);
            shape13.Append(textBody13);

            Shape shape14 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties14 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties18 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties14 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks14 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties14.Append(shapeLocks14);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties18 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape14 = new PlaceholderShape() { Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties18.Append(placeholderShape14);

            nonVisualShapeProperties14.Append(nonVisualDrawingProperties18);
            nonVisualShapeProperties14.Append(nonVisualShapeDrawingProperties14);
            nonVisualShapeProperties14.Append(applicationNonVisualDrawingProperties18);

            ShapeProperties shapeProperties14 = new ShapeProperties();

            A.Transform2D transform2D9 = new A.Transform2D();
            A.Offset offset13 = new A.Offset() { X = 3887391L, Y = 987426L };
            A.Extents extents13 = new A.Extents() { Cx = 4629150L, Cy = 4873625L };

            transform2D9.Append(offset13);
            transform2D9.Append(extents13);

            shapeProperties14.Append(transform2D9);

            TextBody textBody14 = new TextBody();
            A.BodyProperties bodyProperties14 = new A.BodyProperties();

            A.ListStyle listStyle14 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties10 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties35 = new A.DefaultRunProperties() { FontSize = 3200 };

            level1ParagraphProperties10.Append(defaultRunProperties35);

            A.Level2ParagraphProperties level2ParagraphProperties4 = new A.Level2ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties36 = new A.DefaultRunProperties() { FontSize = 2800 };

            level2ParagraphProperties4.Append(defaultRunProperties36);

            A.Level3ParagraphProperties level3ParagraphProperties4 = new A.Level3ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties37 = new A.DefaultRunProperties() { FontSize = 2400 };

            level3ParagraphProperties4.Append(defaultRunProperties37);

            A.Level4ParagraphProperties level4ParagraphProperties4 = new A.Level4ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties38 = new A.DefaultRunProperties() { FontSize = 2000 };

            level4ParagraphProperties4.Append(defaultRunProperties38);

            A.Level5ParagraphProperties level5ParagraphProperties4 = new A.Level5ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties39 = new A.DefaultRunProperties() { FontSize = 2000 };

            level5ParagraphProperties4.Append(defaultRunProperties39);

            A.Level6ParagraphProperties level6ParagraphProperties4 = new A.Level6ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties40 = new A.DefaultRunProperties() { FontSize = 2000 };

            level6ParagraphProperties4.Append(defaultRunProperties40);

            A.Level7ParagraphProperties level7ParagraphProperties4 = new A.Level7ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties41 = new A.DefaultRunProperties() { FontSize = 2000 };

            level7ParagraphProperties4.Append(defaultRunProperties41);

            A.Level8ParagraphProperties level8ParagraphProperties4 = new A.Level8ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties42 = new A.DefaultRunProperties() { FontSize = 2000 };

            level8ParagraphProperties4.Append(defaultRunProperties42);

            A.Level9ParagraphProperties level9ParagraphProperties4 = new A.Level9ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties43 = new A.DefaultRunProperties() { FontSize = 2000 };

            level9ParagraphProperties4.Append(defaultRunProperties43);

            listStyle14.Append(level1ParagraphProperties10);
            listStyle14.Append(level2ParagraphProperties4);
            listStyle14.Append(level3ParagraphProperties4);
            listStyle14.Append(level4ParagraphProperties4);
            listStyle14.Append(level5ParagraphProperties4);
            listStyle14.Append(level6ParagraphProperties4);
            listStyle14.Append(level7ParagraphProperties4);
            listStyle14.Append(level8ParagraphProperties4);
            listStyle14.Append(level9ParagraphProperties4);

            A.Paragraph paragraph18 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties() { Level = 0 };

            A.Run run20 = new A.Run();
            A.RunProperties runProperties24 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text24 = new A.Text();
            text24.Text = "マスター テキストの書式設定";

            run20.Append(runProperties24);
            run20.Append(text24);

            paragraph18.Append(paragraphProperties6);
            paragraph18.Append(run20);

            A.Paragraph paragraph19 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties() { Level = 1 };

            A.Run run21 = new A.Run();
            A.RunProperties runProperties25 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text25 = new A.Text();
            text25.Text = "第 ";

            run21.Append(runProperties25);
            run21.Append(text25);

            A.Run run22 = new A.Run();
            A.RunProperties runProperties26 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text26 = new A.Text();
            text26.Text = "2 ";

            run22.Append(runProperties26);
            run22.Append(text26);

            A.Run run23 = new A.Run();
            A.RunProperties runProperties27 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text27 = new A.Text();
            text27.Text = "レベル";

            run23.Append(runProperties27);
            run23.Append(text27);

            paragraph19.Append(paragraphProperties7);
            paragraph19.Append(run21);
            paragraph19.Append(run22);
            paragraph19.Append(run23);

            A.Paragraph paragraph20 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties() { Level = 2 };

            A.Run run24 = new A.Run();
            A.RunProperties runProperties28 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text28 = new A.Text();
            text28.Text = "第 ";

            run24.Append(runProperties28);
            run24.Append(text28);

            A.Run run25 = new A.Run();
            A.RunProperties runProperties29 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text29 = new A.Text();
            text29.Text = "3 ";

            run25.Append(runProperties29);
            run25.Append(text29);

            A.Run run26 = new A.Run();
            A.RunProperties runProperties30 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text30 = new A.Text();
            text30.Text = "レベル";

            run26.Append(runProperties30);
            run26.Append(text30);

            paragraph20.Append(paragraphProperties8);
            paragraph20.Append(run24);
            paragraph20.Append(run25);
            paragraph20.Append(run26);

            A.Paragraph paragraph21 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties() { Level = 3 };

            A.Run run27 = new A.Run();
            A.RunProperties runProperties31 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text31 = new A.Text();
            text31.Text = "第 ";

            run27.Append(runProperties31);
            run27.Append(text31);

            A.Run run28 = new A.Run();
            A.RunProperties runProperties32 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text32 = new A.Text();
            text32.Text = "4 ";

            run28.Append(runProperties32);
            run28.Append(text32);

            A.Run run29 = new A.Run();
            A.RunProperties runProperties33 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text33 = new A.Text();
            text33.Text = "レベル";

            run29.Append(runProperties33);
            run29.Append(text33);

            paragraph21.Append(paragraphProperties9);
            paragraph21.Append(run27);
            paragraph21.Append(run28);
            paragraph21.Append(run29);

            A.Paragraph paragraph22 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties10 = new A.ParagraphProperties() { Level = 4 };

            A.Run run30 = new A.Run();
            A.RunProperties runProperties34 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text34 = new A.Text();
            text34.Text = "第 ";

            run30.Append(runProperties34);
            run30.Append(text34);

            A.Run run31 = new A.Run();
            A.RunProperties runProperties35 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text35 = new A.Text();
            text35.Text = "5 ";

            run31.Append(runProperties35);
            run31.Append(text35);

            A.Run run32 = new A.Run();
            A.RunProperties runProperties36 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text36 = new A.Text();
            text36.Text = "レベル";

            run32.Append(runProperties36);
            run32.Append(text36);
            A.EndParagraphRunProperties endParagraphRunProperties12 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph22.Append(paragraphProperties10);
            paragraph22.Append(run30);
            paragraph22.Append(run31);
            paragraph22.Append(run32);
            paragraph22.Append(endParagraphRunProperties12);

            textBody14.Append(bodyProperties14);
            textBody14.Append(listStyle14);
            textBody14.Append(paragraph18);
            textBody14.Append(paragraph19);
            textBody14.Append(paragraph20);
            textBody14.Append(paragraph21);
            textBody14.Append(paragraph22);

            shape14.Append(nonVisualShapeProperties14);
            shape14.Append(shapeProperties14);
            shape14.Append(textBody14);

            Shape shape15 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties15 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties19 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties15 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks15 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties15.Append(shapeLocks15);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties19 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape15 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties19.Append(placeholderShape15);

            nonVisualShapeProperties15.Append(nonVisualDrawingProperties19);
            nonVisualShapeProperties15.Append(nonVisualShapeDrawingProperties15);
            nonVisualShapeProperties15.Append(applicationNonVisualDrawingProperties19);

            ShapeProperties shapeProperties15 = new ShapeProperties();

            A.Transform2D transform2D10 = new A.Transform2D();
            A.Offset offset14 = new A.Offset() { X = 629841L, Y = 2057400L };
            A.Extents extents14 = new A.Extents() { Cx = 2949178L, Cy = 3811588L };

            transform2D10.Append(offset14);
            transform2D10.Append(extents14);

            shapeProperties15.Append(transform2D10);

            TextBody textBody15 = new TextBody();
            A.BodyProperties bodyProperties15 = new A.BodyProperties();

            A.ListStyle listStyle15 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties11 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet11 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties44 = new A.DefaultRunProperties() { FontSize = 1600 };

            level1ParagraphProperties11.Append(noBullet11);
            level1ParagraphProperties11.Append(defaultRunProperties44);

            A.Level2ParagraphProperties level2ParagraphProperties5 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet12 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties45 = new A.DefaultRunProperties() { FontSize = 1400 };

            level2ParagraphProperties5.Append(noBullet12);
            level2ParagraphProperties5.Append(defaultRunProperties45);

            A.Level3ParagraphProperties level3ParagraphProperties5 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet13 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties46 = new A.DefaultRunProperties() { FontSize = 1200 };

            level3ParagraphProperties5.Append(noBullet13);
            level3ParagraphProperties5.Append(defaultRunProperties46);

            A.Level4ParagraphProperties level4ParagraphProperties5 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet14 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties47 = new A.DefaultRunProperties() { FontSize = 1000 };

            level4ParagraphProperties5.Append(noBullet14);
            level4ParagraphProperties5.Append(defaultRunProperties47);

            A.Level5ParagraphProperties level5ParagraphProperties5 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet15 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties48 = new A.DefaultRunProperties() { FontSize = 1000 };

            level5ParagraphProperties5.Append(noBullet15);
            level5ParagraphProperties5.Append(defaultRunProperties48);

            A.Level6ParagraphProperties level6ParagraphProperties5 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet16 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties49 = new A.DefaultRunProperties() { FontSize = 1000 };

            level6ParagraphProperties5.Append(noBullet16);
            level6ParagraphProperties5.Append(defaultRunProperties49);

            A.Level7ParagraphProperties level7ParagraphProperties5 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet17 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties50 = new A.DefaultRunProperties() { FontSize = 1000 };

            level7ParagraphProperties5.Append(noBullet17);
            level7ParagraphProperties5.Append(defaultRunProperties50);

            A.Level8ParagraphProperties level8ParagraphProperties5 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet18 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties51 = new A.DefaultRunProperties() { FontSize = 1000 };

            level8ParagraphProperties5.Append(noBullet18);
            level8ParagraphProperties5.Append(defaultRunProperties51);

            A.Level9ParagraphProperties level9ParagraphProperties5 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet19 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties52 = new A.DefaultRunProperties() { FontSize = 1000 };

            level9ParagraphProperties5.Append(noBullet19);
            level9ParagraphProperties5.Append(defaultRunProperties52);

            listStyle15.Append(level1ParagraphProperties11);
            listStyle15.Append(level2ParagraphProperties5);
            listStyle15.Append(level3ParagraphProperties5);
            listStyle15.Append(level4ParagraphProperties5);
            listStyle15.Append(level5ParagraphProperties5);
            listStyle15.Append(level6ParagraphProperties5);
            listStyle15.Append(level7ParagraphProperties5);
            listStyle15.Append(level8ParagraphProperties5);
            listStyle15.Append(level9ParagraphProperties5);

            A.Paragraph paragraph23 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties11 = new A.ParagraphProperties() { Level = 0 };

            A.Run run33 = new A.Run();
            A.RunProperties runProperties37 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text37 = new A.Text();
            text37.Text = "マスター テキストの書式設定";

            run33.Append(runProperties37);
            run33.Append(text37);

            paragraph23.Append(paragraphProperties11);
            paragraph23.Append(run33);

            textBody15.Append(bodyProperties15);
            textBody15.Append(listStyle15);
            textBody15.Append(paragraph23);

            shape15.Append(nonVisualShapeProperties15);
            shape15.Append(shapeProperties15);
            shape15.Append(textBody15);

            Shape shape16 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties16 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties20 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties16 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks16 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties16.Append(shapeLocks16);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties20 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape16 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties20.Append(placeholderShape16);

            nonVisualShapeProperties16.Append(nonVisualDrawingProperties20);
            nonVisualShapeProperties16.Append(nonVisualShapeDrawingProperties16);
            nonVisualShapeProperties16.Append(applicationNonVisualDrawingProperties20);
            ShapeProperties shapeProperties16 = new ShapeProperties();

            TextBody textBody16 = new TextBody();
            A.BodyProperties bodyProperties16 = new A.BodyProperties();
            A.ListStyle listStyle16 = new A.ListStyle();

            A.Paragraph paragraph24 = new A.Paragraph();

            A.Field field5 = new A.Field() { Id = "{BFFF7C5F-97F4-4B42-9C39-61661A529470}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties38 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };
            runProperties38.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text38 = new A.Text();
            text38.Text = "2018/5/3";

            field5.Append(runProperties38);
            field5.Append(text38);
            A.EndParagraphRunProperties endParagraphRunProperties13 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };

            paragraph24.Append(field5);
            paragraph24.Append(endParagraphRunProperties13);

            textBody16.Append(bodyProperties16);
            textBody16.Append(listStyle16);
            textBody16.Append(paragraph24);

            shape16.Append(nonVisualShapeProperties16);
            shape16.Append(shapeProperties16);
            shape16.Append(textBody16);

            Shape shape17 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties17 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties21 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties17 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks17 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties17.Append(shapeLocks17);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties21 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape17 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties21.Append(placeholderShape17);

            nonVisualShapeProperties17.Append(nonVisualDrawingProperties21);
            nonVisualShapeProperties17.Append(nonVisualShapeDrawingProperties17);
            nonVisualShapeProperties17.Append(applicationNonVisualDrawingProperties21);
            ShapeProperties shapeProperties17 = new ShapeProperties();

            TextBody textBody17 = new TextBody();
            A.BodyProperties bodyProperties17 = new A.BodyProperties();
            A.ListStyle listStyle17 = new A.ListStyle();

            A.Paragraph paragraph25 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties14 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };

            paragraph25.Append(endParagraphRunProperties14);

            textBody17.Append(bodyProperties17);
            textBody17.Append(listStyle17);
            textBody17.Append(paragraph25);

            shape17.Append(nonVisualShapeProperties17);
            shape17.Append(shapeProperties17);
            shape17.Append(textBody17);

            Shape shape18 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties18 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties22 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties18 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks18 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties18.Append(shapeLocks18);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties22 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape18 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties22.Append(placeholderShape18);

            nonVisualShapeProperties18.Append(nonVisualDrawingProperties22);
            nonVisualShapeProperties18.Append(nonVisualShapeDrawingProperties18);
            nonVisualShapeProperties18.Append(applicationNonVisualDrawingProperties22);
            ShapeProperties shapeProperties18 = new ShapeProperties();

            TextBody textBody18 = new TextBody();
            A.BodyProperties bodyProperties18 = new A.BodyProperties();
            A.ListStyle listStyle18 = new A.ListStyle();

            A.Paragraph paragraph26 = new A.Paragraph();

            A.Field field6 = new A.Field() { Id = "{10647DAF-1A54-42E0-9176-57F2D0EA6A3B}", Type = "slidenum" };

            A.RunProperties runProperties39 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };
            runProperties39.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text39 = new A.Text();
            text39.Text = "‹#›";

            field6.Append(runProperties39);
            field6.Append(text39);
            A.EndParagraphRunProperties endParagraphRunProperties15 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };

            paragraph26.Append(field6);
            paragraph26.Append(endParagraphRunProperties15);

            textBody18.Append(bodyProperties18);
            textBody18.Append(listStyle18);
            textBody18.Append(paragraph26);

            shape18.Append(nonVisualShapeProperties18);
            shape18.Append(shapeProperties18);
            shape18.Append(textBody18);

            shapeTree4.Append(nonVisualGroupShapeProperties4);
            shapeTree4.Append(groupShapeProperties4);
            shapeTree4.Append(shape13);
            shapeTree4.Append(shape14);
            shapeTree4.Append(shape15);
            shapeTree4.Append(shape16);
            shapeTree4.Append(shape17);
            shapeTree4.Append(shape18);

            CommonSlideDataExtensionList commonSlideDataExtensionList4 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension4 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId4 = new P14.CreationId() { Val = (UInt32Value)539195624U };
            creationId4.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension4.Append(creationId4);

            commonSlideDataExtensionList4.Append(commonSlideDataExtension4);

            commonSlideData4.Append(shapeTree4);
            commonSlideData4.Append(commonSlideDataExtensionList4);

            ColorMapOverride colorMapOverride3 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping3 = new A.MasterColorMapping();

            colorMapOverride3.Append(masterColorMapping3);

            slideLayout2.Append(commonSlideData4);
            slideLayout2.Append(colorMapOverride3);

            slideLayoutPart.SlideLayout = slideLayout2;

            return slideLayoutPart;
        }
    }
}
