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

        private Stack<PPTXFont> FontStack = new Stack<PPTXFont>();

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

            WantReturn = false;
            FontStack.Clear();

            return currentSlide;
        }

        public void Write(PPTXTextRun Text)
        {
            var lastTextArea = AddTextAreaIfEmpty();

            if (lastTextArea.Transform.SizeY != 0)
            {
                AddTextArea();
                lastTextArea = currentSlide.TextAreas.LastOrDefault();
            }

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

            if (FontStack.Count > 0)
            {
                Text.Font = FontStack.Peek();
            }

            lastText.Texts.Add(Text);

            WantReturn = false;
        }

        public void AddTextRow(PPTXText Text)
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
            WantReturn = false;

            var lastTextArea = currentSlide.TextAreas.LastOrDefault();

            currentSlide.TextAreas.Add(new PPTXTextArea(Settings.Margin.Left,
                lastTextArea == null ? 0 : lastTextArea.Transform.PositionY + lastTextArea.Transform.SizeY,
                PageWidth,
                0));
        }

        public void EndTextArea()
        {
            var lastTextArea = AddTextAreaIfEmpty();

            var lastTextAreaSize = 0.0f;

            foreach (var _text in lastTextArea.Texts)
            {
                float maxFontHeight = 0;

                foreach (var _textRun in _text.Texts)
                {
                    maxFontHeight = Math.Max(maxFontHeight, FontHeght(_textRun.Font) + 0.5f);
                }

                lastTextAreaSize += maxFontHeight;
            }

            lastTextArea.Transform.SizeY = lastTextAreaSize;

            WantReturn = false;
        }

        public void PushFont(PPTXFont Font)
        {
            FontStack.Push(Font);
        }

        public void PopFont()
        {
            FontStack.Pop();
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
