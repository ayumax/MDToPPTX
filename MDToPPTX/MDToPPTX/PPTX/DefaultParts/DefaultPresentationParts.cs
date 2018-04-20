using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace MDToPPTX.PPTX.DefaultParts
{
    class DefaultPresentationParts
    {
        public static void CreatePresentationParts(PresentationPart presentationPart)
        {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            SlideSize slideSize1 = new SlideSize()
            {
                Cx = 9144000,
                Cy = 6858000,
                Type = SlideSizeValues.Screen4x3
            };
            NotesSize notesSize1 = new NotesSize()
            {
                Cx = 6858000,
                Cy = 9144000
            };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);


            SlidePart slidePart1 = DefaultSlidePart.CreateSlidePart(presentationPart, "rId2");
            SlideLayoutPart slideLayoutPart1 = DefaultSlideLayoutPart.CreateSlideLayoutPart(slidePart1);
            SlideMasterPart slideMasterPart1 = DefaultSlideMasterPart.CreateSlideMasterPart(slideLayoutPart1);
            ThemePart themePart1 = DefaultTheme.CreateTheme(slideMasterPart1);

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");

        }
    }
}
