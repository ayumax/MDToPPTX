using System;
using System.Collections.Generic;
using System.Text;
using MDToPPTX.PPTX;
using Markdig.Syntax;

namespace MDToPPTX.Markdown.SyntaxWriter
{
    class HeadingBlockWriter : SyntaxWriterBase
    {
        public override void Write(Block Block, SlideManager Slide)
        {
            var headingBlock = Block as HeadingBlock;

            foreach (var line in headingBlock.Inline)
            {
                Slide.currentSlide.TextAreas.Add(new PPTXTextArea(Slide.Settings.Margin.Left, Slide.LastPositionY, 10, 1)
                {
                    Texts = new List<PPTXText>()
                    {
                        new PPTXText(line.ToString(), PPTXBullet.None)
                        {
                            FontSize = 32
                        }
                    }
                });

                Slide.LastPositionY += 2;
            }

        }
    }
}
