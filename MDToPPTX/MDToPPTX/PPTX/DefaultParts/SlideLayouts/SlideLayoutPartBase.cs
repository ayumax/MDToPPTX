using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

namespace MDToPPTX.PPTX.DefaultParts.SlideLayouts
{
    public abstract class SlideLayoutPartBase
    {
        public abstract SlideLayoutPart CreateSlideLayoutPart(OpenXmlPartContainer containerPart, string ID);
    }
}
