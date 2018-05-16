using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MDToPPTX.PPTX;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.SyntaxWriter
{
    class ParagraphBlockWriter : SyntaxWriterBase
    {
        public override void Write(Block Block, SlideManager Slide)
        {
            var headingBlock = Block as ParagraphBlock;

            var textArea = new PPTXTextArea();
            Slide.currentSlide.TextAreas.Add(textArea);


            textArea.Transform = new PPTXTransform(Slide.Settings.Margin.Left, Slide.LastPositionY,
                    PageWidth(Slide), 0);

            bool isConnected = false;

            foreach (var line in headingBlock.Inline)
            {
                if (line is LineBreakInline)
                {
                    var lineBreak = line as LineBreakInline;
                    if (lineBreak.IsHard == false)
                    {
                        isConnected = true;
                    }
                }
                else
                {
                    if (isConnected)
                    {
                        textArea.Texts.Last().Text += " " + line.ToString();
                        isConnected = false;
                    }
                    else
                    {
                        textArea.Texts.Add(new PPTXText(line.ToString(), PPTXBullet.None)
                        {
                            Font = Slide.Settings.NormalFont
                        });

                        textArea.Transform.SizeY += (FontHeght(Slide.Settings.NormalFont.FontSize) + 0.5f);
                    }
                    
                }
                
            }

            Slide.LastPositionY += textArea.Transform.SizeY;


        }
    }
}
