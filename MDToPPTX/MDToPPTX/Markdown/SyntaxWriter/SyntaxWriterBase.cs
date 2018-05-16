using System;
using System.Collections.Generic;
using System.Text;
using MDToPPTX.PPTX;
using Markdig.Syntax;

namespace MDToPPTX.Markdown.SyntaxWriter
{
    abstract class SyntaxWriterBase
    {
        public float FontHeght(float FontSize) => 0.3528f / 10.0f * FontSize;
        public float PageWidth(SlideManager Slide) => Slide.Settings.SlideWidth - (Slide.Settings.Margin.Left + Slide.Settings.Margin.Right);

        public abstract void Write(Block Block, SlideManager Slide);
    }
}
