using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using MDToPPTX.PPTX;

namespace MDToPPTX.Markdown
{
    public class SlideManager
    {
        public PPTXSlide currentSlide { get; private set; }

        public PPTXDocument document { get; private set; }
        public PPTXSetting Settings { get; private set; }

        public float FontHeght(PPTXFont Font) => 0.3528f / 10.0f * Font.FontSize;
        public float PageWidth => Settings.SlideWidth - (Settings.Margin.Left + Settings.Margin.Right);

        private bool WantReturn = false;

        public SlideManager(PPTXDocument document, PPTXSetting Settings)
        {
            this.document = document;
            this.Settings = Settings;

            CreateNewSlide();
        }

        public PPTXSlide CreateNewSlide()
        {
            currentSlide = new PPTXSlide() { SlideLayout = Settings.SlideLayouts[EPPTXSlideLayoutType.BlankSheet] };
            document.Slides.Add(currentSlide);

            return currentSlide;
        }

        public void Write(PPTXTextRun Text)
        {
            var lastTextArea = AddTextAreaIfEmpty();

            if (lastTextArea.Texts.Count == 0)
            {
                lastTextArea.Texts.Add(new PPTXText());
            }

            var lastText = lastTextArea.Texts.Last();
            if (WantReturn)
            {
                lastText = new PPTXText();
                lastTextArea.Texts.Add(lastText);
            }

            lastText.Texts.Add(Text);

            WantReturn = false;
        }

        public void Write(PPTXText Text)
        {
            var lastTextArea = AddTextAreaIfEmpty();

            lastTextArea.Texts.Add(Text);

            WantReturn = false;
        }

        public void WriteReturn()
        {
            WantReturn = true;
        }

        public void AddTextArea()
        {
            var lastTextArea = AddTextAreaIfEmpty();

            currentSlide.TextAreas.Add(new PPTXTextArea(Settings.Margin.Left, lastTextArea.Transform.PositionY + lastTextArea.Transform.SizeY, PageWidth, 0));
        }

        public void EndTextArea()
        {
            var lastTextArea = AddTextAreaIfEmpty();

            var lastTextAreaSize = 0.0f;

            foreach(var _text in lastTextArea.Texts)
            {
                float maxFontHeight = 0;
                
                foreach(var _textRun in _text.Texts)
                {
                    maxFontHeight = Math.Max(maxFontHeight, FontHeght(_textRun.Font) * 1.5f);
                }

                lastTextAreaSize += maxFontHeight;
            }

            lastTextArea.Transform.SizeY = lastTextAreaSize;
        }



        private PPTXTextArea AddTextAreaIfEmpty()
        {
            if (currentSlide.TextAreas.Count == 0)
            {
                currentSlide.TextAreas.Add(new PPTXTextArea(Settings.Margin.Left, Settings.Margin.Top, PageWidth, 0));
            }

            return currentSlide.TextAreas.Last();
        }
    }
}
