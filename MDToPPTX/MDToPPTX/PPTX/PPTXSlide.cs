using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public class PPTXSlide
    {
        public PPTXTextArea Title { get; set; }
        public List<PPTXTextArea> TextAreas { get; set; } = new List<PPTXTextArea>();
        public List<PPTXImage> Images { get; set; } = new List<PPTXImage>();

        public PPTXSlideLayout SlideLayout { get; set; }

    }
}
