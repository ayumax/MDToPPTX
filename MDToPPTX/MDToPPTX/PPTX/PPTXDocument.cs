using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace MDToPPTX.PPTX
{
    public class PPTXDocument : IDisposable
    {
        public List<PPTXSlide> Slides { get; set; }

        private PresentationDocument presentationDoc;

        public PPTXDocument()
        {

        }

        public PPTXDocument(string FilePath, string Title, string SubTitle)
        {
            Init(FilePath, Title, SubTitle);
        }

        public void Init(string FilePath, string Title, string SubTitle)
        {
            presentationDoc = DefaultParts.DefaultPresentationDocument.CreatePresentationDocument(FilePath, Title, SubTitle);
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
