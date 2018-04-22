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

        public PPTXDocument(string FilePath, string Title)
        {
            Init(FilePath, Title);
        }

        public void Init(string FilePath, string Title)
        {
            presentationDoc = DefaultParts.DefaultPresentationDocument.CreatePresentationDocument(FilePath, Title);
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
