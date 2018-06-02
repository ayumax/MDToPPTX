using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace MDToPPTX.PPTX
{
    public class PPTXDocument : IDisposable
    {
        public PPTXSetting FileSettings { get; set; } = new PPTXSetting();
        public PPTXSlideLayoutGroup SlideLayouts { get; private set; } = new PPTXSlideLayoutGroup();
        public List<PPTXSlide> Slides { get; set; } = new List<PPTXSlide>();

        private PresentationDocument presentationDoc;

        public PPTXDocument()
        {

        }

        public PPTXDocument(string FilePath, PPTXSetting FileSettings)
        {
            Init(FilePath, FileSettings);
        }

        public void Init(string FilePath, PPTXSetting FileSettings)
        {
            presentationDoc = DefaultParts.DefaultPresentationDocument.CreatePresentationDocument(FilePath, FileSettings, SlideLayouts);
        }
       
        public void Close()
        {
            foreach(var _slide in Slides)
            {
                OpenXML.SlideWriter writer = new OpenXML.SlideWriter(_slide, SlideLayouts);
                writer.InsertNewSlide(presentationDoc);
            }

            presentationDoc?.Close();
        }

        public void Dispose()
        {
            Close();
        }
        
    }
}
