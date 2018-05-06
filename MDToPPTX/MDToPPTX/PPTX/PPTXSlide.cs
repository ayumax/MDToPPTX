using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public class PPTXSlide
    {
        public PPTXText Title { get; set; }
        public List<PPTXText> Bodys { get; set; } = new List<PPTXText>();
        
        public PPTXSlideLayout SlideLayout { get; set; }

    }
}
