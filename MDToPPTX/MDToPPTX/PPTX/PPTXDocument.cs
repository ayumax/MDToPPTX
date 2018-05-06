using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace MDToPPTX.PPTX
{
    public class PPTXDocument : IDisposable
    {
        public PPTXSetting FileSettings { get; set; } = new PPTXSetting();
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
            presentationDoc = DefaultParts.DefaultPresentationDocument.CreatePresentationDocument(FilePath, FileSettings);
        }
       
        public void Close()
        {
            OpenXML.SlideHelper helper = new OpenXML.SlideHelper();

            foreach(var _slide in Slides)
            {
                helper.InsertNewSlide(presentationDoc, _slide);
            }

            presentationDoc?.Close();
        }

        public void Dispose()
        {
            Close();
        }
        
    }
}
