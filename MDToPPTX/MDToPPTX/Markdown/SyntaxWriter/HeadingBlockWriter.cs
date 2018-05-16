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

            var _font = Slide.Settings.NormalFont;

            switch (headingBlock.Level)
            {
                case 1:
                    _font = Slide.Settings.Header1font;
                    break;
                case 2:
                    _font = Slide.Settings.Header2font;
                    break;
            }

            foreach (var line in headingBlock.Inline)
            {
                Slide.currentSlide.TextAreas.Add(new PPTXTextArea(Slide.Settings.Margin.Left, Slide.LastPositionY, PageWidth(Slide), FontHeght(_font.FontSize))
                {
                    Texts = new List<PPTXText>()
                    {
                        new PPTXText(line.ToString(), PPTXBullet.None)
                        {
                            Font = _font
                        }
                    }
                });

                Slide.LastPositionY += FontHeght(_font.FontSize);
            }

        }
    }
}
