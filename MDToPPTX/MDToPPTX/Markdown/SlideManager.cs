using System;
using System.Collections.Generic;
using System.Text;
using MDToPPTX.PPTX;

namespace MDToPPTX.Markdown
{
    class SlideManager
    {
        public PPTXSlide currentSlide { get; private set; }
        public float LastPositionY { get; set; }

        public PPTXDocument document { get; private set; }
        public PPTXSetting Settings { get; private set; }

        public SlideManager(PPTXDocument document, PPTXSetting Settings)
        {
            LastPositionY = 0;

            this.document = document;
            this.Settings = Settings;

            CreateNewSlide();
        }

        public PPTXSlide CreateNewSlide()
        {
            currentSlide = new PPTXSlide() { SlideLayout = Settings.SlideLayouts[EPPTXSlideLayoutType.BlankSheet] };
            document.Slides.Add(currentSlide);

            LastPositionY = Settings.Margin.Top;

            return currentSlide;
        }
    }
}
