using System;
using System.Collections.Generic;
using System.Text;
using MDToPPTX.PPTX;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.SyntaxWriter
{
    class FencedCodeBlockWriter : SyntaxWriterBase
    {
        public override void Write(Block Block, SlideManager Slide)
        {
            var codeBlock = Block as FencedCodeBlock;

            var textArea = new PPTXTextArea();
            Slide.currentSlide.TextAreas.Add(textArea);


            textArea.Transform = new PPTXTransform(Slide.Settings.Margin.Left, Slide.LastPositionY,
                    PageWidth(Slide), 0);

            var codeString = "";

            foreach (var line in codeBlock.Lines)
            {
                codeString += line + "\n";
            }

            textArea.Texts.Add(new PPTXText(codeString, PPTXBullet.None)
            {
                Font = Slide.Settings.CodeFont
            });

            textArea.Transform.SizeY += (FontHeght(Slide.Settings.CodeFont.FontSize) + 0.5f) * codeBlock.Lines.Count;
            Slide.LastPositionY += textArea.Transform.SizeY;


        }
    }
}
