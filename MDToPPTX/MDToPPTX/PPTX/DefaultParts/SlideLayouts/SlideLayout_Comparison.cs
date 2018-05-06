using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;


namespace MDToPPTX.PPTX.DefaultParts.SlideLayouts
{
    public class SlideLayout_Comparison : SlideLayoutPartBase
    {
        public override SlideLayoutPart CreateSlideLayoutPart(OpenXmlPartContainer containerPart)
        {
            SlideLayoutPart slideLayoutPart = containerPart.AddNewPart<SlideLayoutPart>(LayoutSetting.ID);

            SlideLayout slideLayout8 = new SlideLayout() { Type = SlideLayoutValues.TwoTextAndTwoObjects, Preserve = true };
            slideLayout8.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout8.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout8.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData10 = new CommonSlideData() { Name = LayoutSetting.Name };

            ShapeTree shapeTree10 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties10 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties50 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties10 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties50 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties10.Append(nonVisualDrawingProperties50);
            nonVisualGroupShapeProperties10.Append(nonVisualGroupShapeDrawingProperties10);
            nonVisualGroupShapeProperties10.Append(applicationNonVisualDrawingProperties50);

            GroupShapeProperties groupShapeProperties10 = new GroupShapeProperties();

            A.TransformGroup transformGroup10 = new A.TransformGroup();
            A.Offset offset24 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents24 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset10 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents10 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup10.Append(offset24);
            transformGroup10.Append(extents24);
            transformGroup10.Append(childOffset10);
            transformGroup10.Append(childExtents10);

            groupShapeProperties10.Append(transformGroup10);

            Shape shape41 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties41 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties51 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties41 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks41 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties41.Append(shapeLocks41);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties51 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape41 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties51.Append(placeholderShape41);

            nonVisualShapeProperties41.Append(nonVisualDrawingProperties51);
            nonVisualShapeProperties41.Append(nonVisualShapeDrawingProperties41);
            nonVisualShapeProperties41.Append(applicationNonVisualDrawingProperties51);

            ShapeProperties shapeProperties41 = new ShapeProperties();

            A.Transform2D transform2D15 = new A.Transform2D();
            A.Offset offset25 = new A.Offset() { X = 629841L, Y = 365126L };
            A.Extents extents25 = new A.Extents() { Cx = 7886700L, Cy = 1325563L };

            transform2D15.Append(offset25);
            transform2D15.Append(extents25);

            shapeProperties41.Append(transform2D15);

            TextBody textBody41 = new TextBody();
            A.BodyProperties bodyProperties41 = new A.BodyProperties();
            A.ListStyle listStyle41 = new A.ListStyle();

            A.Paragraph paragraph57 = new A.Paragraph();

            A.Run run65 = new A.Run();
            A.RunProperties runProperties81 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text81 = new A.Text();
            text81.Text = "マスター タイトルの書式設定";

            run65.Append(runProperties81);
            run65.Append(text81);
            A.EndParagraphRunProperties endParagraphRunProperties37 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph57.Append(run65);
            paragraph57.Append(endParagraphRunProperties37);

            textBody41.Append(bodyProperties41);
            textBody41.Append(listStyle41);
            textBody41.Append(paragraph57);

            shape41.Append(nonVisualShapeProperties41);
            shape41.Append(shapeProperties41);
            shape41.Append(textBody41);

            Shape shape42 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties42 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties52 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties42 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks42 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties42.Append(shapeLocks42);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties52 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape42 = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties52.Append(placeholderShape42);

            nonVisualShapeProperties42.Append(nonVisualDrawingProperties52);
            nonVisualShapeProperties42.Append(nonVisualShapeDrawingProperties42);
            nonVisualShapeProperties42.Append(applicationNonVisualDrawingProperties52);

            ShapeProperties shapeProperties42 = new ShapeProperties();

            A.Transform2D transform2D16 = new A.Transform2D();
            A.Offset offset26 = new A.Offset() { X = 629842L, Y = 1681163L };
            A.Extents extents26 = new A.Extents() { Cx = 3868340L, Cy = 823912L };

            transform2D16.Append(offset26);
            transform2D16.Append(extents26);

            shapeProperties42.Append(transform2D16);

            TextBody textBody42 = new TextBody();
            A.BodyProperties bodyProperties42 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle42 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties14 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet29 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties63 = new A.DefaultRunProperties() { FontSize = 2400, Bold = true };

            level1ParagraphProperties14.Append(noBullet29);
            level1ParagraphProperties14.Append(defaultRunProperties63);

            A.Level2ParagraphProperties level2ParagraphProperties7 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet30 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties64 = new A.DefaultRunProperties() { FontSize = 2000, Bold = true };

            level2ParagraphProperties7.Append(noBullet30);
            level2ParagraphProperties7.Append(defaultRunProperties64);

            A.Level3ParagraphProperties level3ParagraphProperties7 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet31 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties65 = new A.DefaultRunProperties() { FontSize = 1800, Bold = true };

            level3ParagraphProperties7.Append(noBullet31);
            level3ParagraphProperties7.Append(defaultRunProperties65);

            A.Level4ParagraphProperties level4ParagraphProperties7 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet32 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties66 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level4ParagraphProperties7.Append(noBullet32);
            level4ParagraphProperties7.Append(defaultRunProperties66);

            A.Level5ParagraphProperties level5ParagraphProperties7 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet33 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties67 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level5ParagraphProperties7.Append(noBullet33);
            level5ParagraphProperties7.Append(defaultRunProperties67);

            A.Level6ParagraphProperties level6ParagraphProperties7 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet34 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties68 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level6ParagraphProperties7.Append(noBullet34);
            level6ParagraphProperties7.Append(defaultRunProperties68);

            A.Level7ParagraphProperties level7ParagraphProperties7 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet35 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties69 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level7ParagraphProperties7.Append(noBullet35);
            level7ParagraphProperties7.Append(defaultRunProperties69);

            A.Level8ParagraphProperties level8ParagraphProperties7 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet36 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties70 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level8ParagraphProperties7.Append(noBullet36);
            level8ParagraphProperties7.Append(defaultRunProperties70);

            A.Level9ParagraphProperties level9ParagraphProperties7 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet37 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties71 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level9ParagraphProperties7.Append(noBullet37);
            level9ParagraphProperties7.Append(defaultRunProperties71);

            listStyle42.Append(level1ParagraphProperties14);
            listStyle42.Append(level2ParagraphProperties7);
            listStyle42.Append(level3ParagraphProperties7);
            listStyle42.Append(level4ParagraphProperties7);
            listStyle42.Append(level5ParagraphProperties7);
            listStyle42.Append(level6ParagraphProperties7);
            listStyle42.Append(level7ParagraphProperties7);
            listStyle42.Append(level8ParagraphProperties7);
            listStyle42.Append(level9ParagraphProperties7);

            A.Paragraph paragraph58 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties23 = new A.ParagraphProperties() { Level = 0 };

            A.Run run66 = new A.Run();
            A.RunProperties runProperties82 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text82 = new A.Text();
            text82.Text = "マスター テキストの書式設定";

            run66.Append(runProperties82);
            run66.Append(text82);

            paragraph58.Append(paragraphProperties23);
            paragraph58.Append(run66);

            textBody42.Append(bodyProperties42);
            textBody42.Append(listStyle42);
            textBody42.Append(paragraph58);

            shape42.Append(nonVisualShapeProperties42);
            shape42.Append(shapeProperties42);
            shape42.Append(textBody42);

            Shape shape43 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties43 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties53 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties43 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks43 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties43.Append(shapeLocks43);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties53 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape43 = new PlaceholderShape() { Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties53.Append(placeholderShape43);

            nonVisualShapeProperties43.Append(nonVisualDrawingProperties53);
            nonVisualShapeProperties43.Append(nonVisualShapeDrawingProperties43);
            nonVisualShapeProperties43.Append(applicationNonVisualDrawingProperties53);

            ShapeProperties shapeProperties43 = new ShapeProperties();

            A.Transform2D transform2D17 = new A.Transform2D();
            A.Offset offset27 = new A.Offset() { X = 629842L, Y = 2505075L };
            A.Extents extents27 = new A.Extents() { Cx = 3868340L, Cy = 3684588L };

            transform2D17.Append(offset27);
            transform2D17.Append(extents27);

            shapeProperties43.Append(transform2D17);

            TextBody textBody43 = new TextBody();
            A.BodyProperties bodyProperties43 = new A.BodyProperties();
            A.ListStyle listStyle43 = new A.ListStyle();

            A.Paragraph paragraph59 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties24 = new A.ParagraphProperties() { Level = 0 };

            A.Run run67 = new A.Run();
            A.RunProperties runProperties83 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text83 = new A.Text();
            text83.Text = "マスター テキストの書式設定";

            run67.Append(runProperties83);
            run67.Append(text83);

            paragraph59.Append(paragraphProperties24);
            paragraph59.Append(run67);

            A.Paragraph paragraph60 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties25 = new A.ParagraphProperties() { Level = 1 };

            A.Run run68 = new A.Run();
            A.RunProperties runProperties84 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text84 = new A.Text();
            text84.Text = "第 ";

            run68.Append(runProperties84);
            run68.Append(text84);

            A.Run run69 = new A.Run();
            A.RunProperties runProperties85 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text85 = new A.Text();
            text85.Text = "2 ";

            run69.Append(runProperties85);
            run69.Append(text85);

            A.Run run70 = new A.Run();
            A.RunProperties runProperties86 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text86 = new A.Text();
            text86.Text = "レベル";

            run70.Append(runProperties86);
            run70.Append(text86);

            paragraph60.Append(paragraphProperties25);
            paragraph60.Append(run68);
            paragraph60.Append(run69);
            paragraph60.Append(run70);

            A.Paragraph paragraph61 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties26 = new A.ParagraphProperties() { Level = 2 };

            A.Run run71 = new A.Run();
            A.RunProperties runProperties87 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text87 = new A.Text();
            text87.Text = "第 ";

            run71.Append(runProperties87);
            run71.Append(text87);

            A.Run run72 = new A.Run();
            A.RunProperties runProperties88 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text88 = new A.Text();
            text88.Text = "3 ";

            run72.Append(runProperties88);
            run72.Append(text88);

            A.Run run73 = new A.Run();
            A.RunProperties runProperties89 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text89 = new A.Text();
            text89.Text = "レベル";

            run73.Append(runProperties89);
            run73.Append(text89);

            paragraph61.Append(paragraphProperties26);
            paragraph61.Append(run71);
            paragraph61.Append(run72);
            paragraph61.Append(run73);

            A.Paragraph paragraph62 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties27 = new A.ParagraphProperties() { Level = 3 };

            A.Run run74 = new A.Run();
            A.RunProperties runProperties90 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text90 = new A.Text();
            text90.Text = "第 ";

            run74.Append(runProperties90);
            run74.Append(text90);

            A.Run run75 = new A.Run();
            A.RunProperties runProperties91 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text91 = new A.Text();
            text91.Text = "4 ";

            run75.Append(runProperties91);
            run75.Append(text91);

            A.Run run76 = new A.Run();
            A.RunProperties runProperties92 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text92 = new A.Text();
            text92.Text = "レベル";

            run76.Append(runProperties92);
            run76.Append(text92);

            paragraph62.Append(paragraphProperties27);
            paragraph62.Append(run74);
            paragraph62.Append(run75);
            paragraph62.Append(run76);

            A.Paragraph paragraph63 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties28 = new A.ParagraphProperties() { Level = 4 };

            A.Run run77 = new A.Run();
            A.RunProperties runProperties93 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text93 = new A.Text();
            text93.Text = "第 ";

            run77.Append(runProperties93);
            run77.Append(text93);

            A.Run run78 = new A.Run();
            A.RunProperties runProperties94 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text94 = new A.Text();
            text94.Text = "5 ";

            run78.Append(runProperties94);
            run78.Append(text94);

            A.Run run79 = new A.Run();
            A.RunProperties runProperties95 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text95 = new A.Text();
            text95.Text = "レベル";

            run79.Append(runProperties95);
            run79.Append(text95);
            A.EndParagraphRunProperties endParagraphRunProperties38 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph63.Append(paragraphProperties28);
            paragraph63.Append(run77);
            paragraph63.Append(run78);
            paragraph63.Append(run79);
            paragraph63.Append(endParagraphRunProperties38);

            textBody43.Append(bodyProperties43);
            textBody43.Append(listStyle43);
            textBody43.Append(paragraph59);
            textBody43.Append(paragraph60);
            textBody43.Append(paragraph61);
            textBody43.Append(paragraph62);
            textBody43.Append(paragraph63);

            shape43.Append(nonVisualShapeProperties43);
            shape43.Append(shapeProperties43);
            shape43.Append(textBody43);

            Shape shape44 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties44 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties54 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Text Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties44 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks44 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties44.Append(shapeLocks44);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties54 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape44 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties54.Append(placeholderShape44);

            nonVisualShapeProperties44.Append(nonVisualDrawingProperties54);
            nonVisualShapeProperties44.Append(nonVisualShapeDrawingProperties44);
            nonVisualShapeProperties44.Append(applicationNonVisualDrawingProperties54);

            ShapeProperties shapeProperties44 = new ShapeProperties();

            A.Transform2D transform2D18 = new A.Transform2D();
            A.Offset offset28 = new A.Offset() { X = 4629150L, Y = 1681163L };
            A.Extents extents28 = new A.Extents() { Cx = 3887391L, Cy = 823912L };

            transform2D18.Append(offset28);
            transform2D18.Append(extents28);

            shapeProperties44.Append(transform2D18);

            TextBody textBody44 = new TextBody();
            A.BodyProperties bodyProperties44 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle44 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties15 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet38 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties72 = new A.DefaultRunProperties() { FontSize = 2400, Bold = true };

            level1ParagraphProperties15.Append(noBullet38);
            level1ParagraphProperties15.Append(defaultRunProperties72);

            A.Level2ParagraphProperties level2ParagraphProperties8 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet39 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties73 = new A.DefaultRunProperties() { FontSize = 2000, Bold = true };

            level2ParagraphProperties8.Append(noBullet39);
            level2ParagraphProperties8.Append(defaultRunProperties73);

            A.Level3ParagraphProperties level3ParagraphProperties8 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet40 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties74 = new A.DefaultRunProperties() { FontSize = 1800, Bold = true };

            level3ParagraphProperties8.Append(noBullet40);
            level3ParagraphProperties8.Append(defaultRunProperties74);

            A.Level4ParagraphProperties level4ParagraphProperties8 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet41 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties75 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level4ParagraphProperties8.Append(noBullet41);
            level4ParagraphProperties8.Append(defaultRunProperties75);

            A.Level5ParagraphProperties level5ParagraphProperties8 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet42 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties76 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level5ParagraphProperties8.Append(noBullet42);
            level5ParagraphProperties8.Append(defaultRunProperties76);

            A.Level6ParagraphProperties level6ParagraphProperties8 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet43 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties77 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level6ParagraphProperties8.Append(noBullet43);
            level6ParagraphProperties8.Append(defaultRunProperties77);

            A.Level7ParagraphProperties level7ParagraphProperties8 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet44 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties78 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level7ParagraphProperties8.Append(noBullet44);
            level7ParagraphProperties8.Append(defaultRunProperties78);

            A.Level8ParagraphProperties level8ParagraphProperties8 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet45 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties79 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level8ParagraphProperties8.Append(noBullet45);
            level8ParagraphProperties8.Append(defaultRunProperties79);

            A.Level9ParagraphProperties level9ParagraphProperties8 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet46 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties80 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level9ParagraphProperties8.Append(noBullet46);
            level9ParagraphProperties8.Append(defaultRunProperties80);

            listStyle44.Append(level1ParagraphProperties15);
            listStyle44.Append(level2ParagraphProperties8);
            listStyle44.Append(level3ParagraphProperties8);
            listStyle44.Append(level4ParagraphProperties8);
            listStyle44.Append(level5ParagraphProperties8);
            listStyle44.Append(level6ParagraphProperties8);
            listStyle44.Append(level7ParagraphProperties8);
            listStyle44.Append(level8ParagraphProperties8);
            listStyle44.Append(level9ParagraphProperties8);

            A.Paragraph paragraph64 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties29 = new A.ParagraphProperties() { Level = 0 };

            A.Run run80 = new A.Run();
            A.RunProperties runProperties96 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text96 = new A.Text();
            text96.Text = "マスター テキストの書式設定";

            run80.Append(runProperties96);
            run80.Append(text96);

            paragraph64.Append(paragraphProperties29);
            paragraph64.Append(run80);

            textBody44.Append(bodyProperties44);
            textBody44.Append(listStyle44);
            textBody44.Append(paragraph64);

            shape44.Append(nonVisualShapeProperties44);
            shape44.Append(shapeProperties44);
            shape44.Append(textBody44);

            Shape shape45 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties45 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties55 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Content Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties45 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks45 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties45.Append(shapeLocks45);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties55 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape45 = new PlaceholderShape() { Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties55.Append(placeholderShape45);

            nonVisualShapeProperties45.Append(nonVisualDrawingProperties55);
            nonVisualShapeProperties45.Append(nonVisualShapeDrawingProperties45);
            nonVisualShapeProperties45.Append(applicationNonVisualDrawingProperties55);

            ShapeProperties shapeProperties45 = new ShapeProperties();

            A.Transform2D transform2D19 = new A.Transform2D();
            A.Offset offset29 = new A.Offset() { X = 4629150L, Y = 2505075L };
            A.Extents extents29 = new A.Extents() { Cx = 3887391L, Cy = 3684588L };

            transform2D19.Append(offset29);
            transform2D19.Append(extents29);

            shapeProperties45.Append(transform2D19);

            TextBody textBody45 = new TextBody();
            A.BodyProperties bodyProperties45 = new A.BodyProperties();
            A.ListStyle listStyle45 = new A.ListStyle();

            A.Paragraph paragraph65 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties30 = new A.ParagraphProperties() { Level = 0 };

            A.Run run81 = new A.Run();
            A.RunProperties runProperties97 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text97 = new A.Text();
            text97.Text = "マスター テキストの書式設定";

            run81.Append(runProperties97);
            run81.Append(text97);

            paragraph65.Append(paragraphProperties30);
            paragraph65.Append(run81);

            A.Paragraph paragraph66 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties31 = new A.ParagraphProperties() { Level = 1 };

            A.Run run82 = new A.Run();
            A.RunProperties runProperties98 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text98 = new A.Text();
            text98.Text = "第 ";

            run82.Append(runProperties98);
            run82.Append(text98);

            A.Run run83 = new A.Run();
            A.RunProperties runProperties99 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text99 = new A.Text();
            text99.Text = "2 ";

            run83.Append(runProperties99);
            run83.Append(text99);

            A.Run run84 = new A.Run();
            A.RunProperties runProperties100 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text100 = new A.Text();
            text100.Text = "レベル";

            run84.Append(runProperties100);
            run84.Append(text100);

            paragraph66.Append(paragraphProperties31);
            paragraph66.Append(run82);
            paragraph66.Append(run83);
            paragraph66.Append(run84);

            A.Paragraph paragraph67 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties32 = new A.ParagraphProperties() { Level = 2 };

            A.Run run85 = new A.Run();
            A.RunProperties runProperties101 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text101 = new A.Text();
            text101.Text = "第 ";

            run85.Append(runProperties101);
            run85.Append(text101);

            A.Run run86 = new A.Run();
            A.RunProperties runProperties102 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text102 = new A.Text();
            text102.Text = "3 ";

            run86.Append(runProperties102);
            run86.Append(text102);

            A.Run run87 = new A.Run();
            A.RunProperties runProperties103 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text103 = new A.Text();
            text103.Text = "レベル";

            run87.Append(runProperties103);
            run87.Append(text103);

            paragraph67.Append(paragraphProperties32);
            paragraph67.Append(run85);
            paragraph67.Append(run86);
            paragraph67.Append(run87);

            A.Paragraph paragraph68 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties33 = new A.ParagraphProperties() { Level = 3 };

            A.Run run88 = new A.Run();
            A.RunProperties runProperties104 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text104 = new A.Text();
            text104.Text = "第 ";

            run88.Append(runProperties104);
            run88.Append(text104);

            A.Run run89 = new A.Run();
            A.RunProperties runProperties105 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text105 = new A.Text();
            text105.Text = "4 ";

            run89.Append(runProperties105);
            run89.Append(text105);

            A.Run run90 = new A.Run();
            A.RunProperties runProperties106 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text106 = new A.Text();
            text106.Text = "レベル";

            run90.Append(runProperties106);
            run90.Append(text106);

            paragraph68.Append(paragraphProperties33);
            paragraph68.Append(run88);
            paragraph68.Append(run89);
            paragraph68.Append(run90);

            A.Paragraph paragraph69 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties34 = new A.ParagraphProperties() { Level = 4 };

            A.Run run91 = new A.Run();
            A.RunProperties runProperties107 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text107 = new A.Text();
            text107.Text = "第 ";

            run91.Append(runProperties107);
            run91.Append(text107);

            A.Run run92 = new A.Run();
            A.RunProperties runProperties108 = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "ja-JP" };
            A.Text text108 = new A.Text();
            text108.Text = "5 ";

            run92.Append(runProperties108);
            run92.Append(text108);

            A.Run run93 = new A.Run();
            A.RunProperties runProperties109 = new A.RunProperties() { Language = "ja-JP", AlternativeLanguage = "en-US" };
            A.Text text109 = new A.Text();
            text109.Text = "レベル";

            run93.Append(runProperties109);
            run93.Append(text109);
            A.EndParagraphRunProperties endParagraphRunProperties39 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph69.Append(paragraphProperties34);
            paragraph69.Append(run91);
            paragraph69.Append(run92);
            paragraph69.Append(run93);
            paragraph69.Append(endParagraphRunProperties39);

            textBody45.Append(bodyProperties45);
            textBody45.Append(listStyle45);
            textBody45.Append(paragraph65);
            textBody45.Append(paragraph66);
            textBody45.Append(paragraph67);
            textBody45.Append(paragraph68);
            textBody45.Append(paragraph69);

            shape45.Append(nonVisualShapeProperties45);
            shape45.Append(shapeProperties45);
            shape45.Append(textBody45);

            Shape shape46 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties46 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties56 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Date Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties46 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks46 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties46.Append(shapeLocks46);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties56 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape46 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties56.Append(placeholderShape46);

            nonVisualShapeProperties46.Append(nonVisualDrawingProperties56);
            nonVisualShapeProperties46.Append(nonVisualShapeDrawingProperties46);
            nonVisualShapeProperties46.Append(applicationNonVisualDrawingProperties56);
            ShapeProperties shapeProperties46 = new ShapeProperties();

            TextBody textBody46 = new TextBody();
            A.BodyProperties bodyProperties46 = new A.BodyProperties();
            A.ListStyle listStyle46 = new A.ListStyle();

            A.Paragraph paragraph70 = new A.Paragraph();

            A.Field field17 = new A.Field() { Id = "{BFFF7C5F-97F4-4B42-9C39-61661A529470}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties110 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };
            runProperties110.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text110 = new A.Text();
            text110.Text = "2018/5/3";

            field17.Append(runProperties110);
            field17.Append(text110);
            A.EndParagraphRunProperties endParagraphRunProperties40 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };

            paragraph70.Append(field17);
            paragraph70.Append(endParagraphRunProperties40);

            textBody46.Append(bodyProperties46);
            textBody46.Append(listStyle46);
            textBody46.Append(paragraph70);

            shape46.Append(nonVisualShapeProperties46);
            shape46.Append(shapeProperties46);
            shape46.Append(textBody46);

            Shape shape47 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties47 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties57 = new NonVisualDrawingProperties() { Id = (UInt32Value)8U, Name = "Footer Placeholder 7" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties47 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks47 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties47.Append(shapeLocks47);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties57 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape47 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties57.Append(placeholderShape47);

            nonVisualShapeProperties47.Append(nonVisualDrawingProperties57);
            nonVisualShapeProperties47.Append(nonVisualShapeDrawingProperties47);
            nonVisualShapeProperties47.Append(applicationNonVisualDrawingProperties57);
            ShapeProperties shapeProperties47 = new ShapeProperties();

            TextBody textBody47 = new TextBody();
            A.BodyProperties bodyProperties47 = new A.BodyProperties();
            A.ListStyle listStyle47 = new A.ListStyle();

            A.Paragraph paragraph71 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties41 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };

            paragraph71.Append(endParagraphRunProperties41);

            textBody47.Append(bodyProperties47);
            textBody47.Append(listStyle47);
            textBody47.Append(paragraph71);

            shape47.Append(nonVisualShapeProperties47);
            shape47.Append(shapeProperties47);
            shape47.Append(textBody47);

            Shape shape48 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties48 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties58 = new NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Slide Number Placeholder 8" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties48 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks48 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties48.Append(shapeLocks48);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties58 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape48 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties58.Append(placeholderShape48);

            nonVisualShapeProperties48.Append(nonVisualDrawingProperties58);
            nonVisualShapeProperties48.Append(nonVisualShapeDrawingProperties48);
            nonVisualShapeProperties48.Append(applicationNonVisualDrawingProperties58);
            ShapeProperties shapeProperties48 = new ShapeProperties();

            TextBody textBody48 = new TextBody();
            A.BodyProperties bodyProperties48 = new A.BodyProperties();
            A.ListStyle listStyle48 = new A.ListStyle();

            A.Paragraph paragraph72 = new A.Paragraph();

            A.Field field18 = new A.Field() { Id = "{10647DAF-1A54-42E0-9176-57F2D0EA6A3B}", Type = "slidenum" };

            A.RunProperties runProperties111 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };
            runProperties111.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text111 = new A.Text();
            text111.Text = "‹#›";

            field18.Append(runProperties111);
            field18.Append(text111);
            A.EndParagraphRunProperties endParagraphRunProperties42 = new A.EndParagraphRunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US" };

            paragraph72.Append(field18);
            paragraph72.Append(endParagraphRunProperties42);

            textBody48.Append(bodyProperties48);
            textBody48.Append(listStyle48);
            textBody48.Append(paragraph72);

            shape48.Append(nonVisualShapeProperties48);
            shape48.Append(shapeProperties48);
            shape48.Append(textBody48);

            shapeTree10.Append(nonVisualGroupShapeProperties10);
            shapeTree10.Append(groupShapeProperties10);
            shapeTree10.Append(shape41);
            shapeTree10.Append(shape42);
            shapeTree10.Append(shape43);
            shapeTree10.Append(shape44);
            shapeTree10.Append(shape45);
            shapeTree10.Append(shape46);
            shapeTree10.Append(shape47);
            shapeTree10.Append(shape48);

            CommonSlideDataExtensionList commonSlideDataExtensionList10 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension10 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId10 = new P14.CreationId() { Val = (UInt32Value)2181463977U };
            creationId10.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension10.Append(creationId10);

            commonSlideDataExtensionList10.Append(commonSlideDataExtension10);

            commonSlideData10.Append(shapeTree10);
            commonSlideData10.Append(commonSlideDataExtensionList10);

            ColorMapOverride colorMapOverride9 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping9 = new A.MasterColorMapping();

            colorMapOverride9.Append(masterColorMapping9);

            slideLayout8.Append(commonSlideData10);
            slideLayout8.Append(colorMapOverride9);

            slideLayoutPart.SlideLayout = slideLayout8;

            return slideLayoutPart;
        }
    }
}
