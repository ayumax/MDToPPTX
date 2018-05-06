using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using MDToPPTX.PPTX.DefaultParts.SlideLayouts;

namespace MDToPPTX.PPTX.DefaultParts
{
    internal class DefaultPresentationDocument
    {
        public static PresentationDocument CreatePresentationDocument(string FilePath, PPTXSetting FileSettings)
        {
            var presentationDoc = PresentationDocument.Create(FilePath, PresentationDocumentType.Presentation);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            DefaultPresentationPart.CreatePresentationPart(presentationPart, FileSettings);

            return presentationDoc;
        }
    }
}
