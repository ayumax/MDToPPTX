using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace MDToPPTX.PPTX
{
    public class PPTXDocument
    {
        public PPTXSetting FileSettings { get; set; } = new PPTXSetting();
        public PPTXSlideLayoutGroup SlideLayouts { get; private set; } = new PPTXSlideLayoutGroup();
        public List<PPTXSlide> Slides { get; set; } = new List<PPTXSlide>();

        public PPTXDocument()
        {

        }
       
        public void SaveAs(string FilePath, PPTXSetting FileSettings)
        {
            var presentationDoc = DefaultParts.DefaultPresentationDocument.CreatePresentationDocument(FilePath, FileSettings, SlideLayouts);

            foreach (var _slide in Slides)
            {
                OpenXML.SlideWriter writer = new OpenXML.SlideWriter(_slide, SlideLayouts);
                writer.InsertNewSlide(presentationDoc);
            }

            presentationDoc?.Close();
        }
        
    }
}
