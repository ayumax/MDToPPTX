using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public class PPTXSlide
    {
        public PPTXText Title { get; set; }
        public List<PPTXText> Texts { get; set; } = new List<PPTXText>();
        public List<PPTXImage> Images { get; set; } = new List<PPTXImage>();

        public PPTXSlideLayout SlideLayout { get; set; }

    }
}
