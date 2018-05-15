using System;
using System.Collections.Generic;
using System.Text;
using MDToPPTX.PPTX;
using Markdig.Syntax;

namespace MDToPPTX.Markdown.SyntaxWriter
{
    class ParagraphBlockWriter : SyntaxWriterBase
    {
        public override void Write(Block Block, SlideManager Slide)
        {
            var headingBlock = Block as ParagraphBlock;

            foreach (var line in headingBlock.Inline)
            {
                Slide.currentSlide.TextAreas.Add(new PPTXTextArea(line.ToString(), 0, Slide.LastPositionY, 10, 1));

                Slide.LastPositionY += 2;
            }

        }
    }
}
