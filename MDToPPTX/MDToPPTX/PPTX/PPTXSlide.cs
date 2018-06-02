using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public class PPTXSlide
    {
        public List<PPTXTextArea> TextAreas { get; set; } = new List<PPTXTextArea>();
        public List<PPTXImage> Images { get; set; } = new List<PPTXImage>();
        public List<PPTXTable> Tables { get; set; } = new List<PPTXTable>();

        // 現状Blanksheetのみ対応
        public EPPTXSlideLayoutType SlideLayout { get; set; } = EPPTXSlideLayoutType.BlankSheet;

    }
}
