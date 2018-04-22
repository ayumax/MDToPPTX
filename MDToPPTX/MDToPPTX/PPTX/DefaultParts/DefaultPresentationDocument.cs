using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace MDToPPTX.PPTX.DefaultParts
{
    internal class DefaultPresentationDocument
    {
        public static PresentationDocument CreatePresentationDocument(string FilePath, string Title)
        {
            var presentationDoc = PresentationDocument.Create(FilePath, PresentationDocumentType.Presentation);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            DefaultPresentationParts.CreatePresentationParts(presentationPart, Title);

            return presentationDoc;
        }
    }
}
