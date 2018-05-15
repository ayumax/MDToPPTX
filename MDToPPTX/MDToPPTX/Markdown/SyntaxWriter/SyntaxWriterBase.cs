using System;
using System.Collections.Generic;
using System.Text;
using MDToPPTX.PPTX;
using Markdig.Syntax;

namespace MDToPPTX.Markdown.SyntaxWriter
{
    abstract class SyntaxWriterBase
    {
        public abstract void Write(Block Block, SlideManager Slide);
    }
}
