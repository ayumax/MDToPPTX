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

        public float FontHeght(PPTXFont Font) => 0.35278f / 10.0f * Font.FontSize;
        public float PageWidth => Settings.SlideWidth - (Settings.Margin.Left + Settings.Margin.Right);

        private bool WantReturn = false;

        private Stack<PPTXFont> FontStack = new Stack<PPTXFont>();
        private Stack<PPTXLink> LinkStack = new Stack<PPTXLink>();

        private PPTXTransform LastAddedItemTransform = new PPTXTransform();

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
            LastAddedItemTransform = new PPTXTransform();

            return currentSlide;
        }

        public void Write(PPTXTextRun Text)
        {
            var lastTextArea = AddTextAreaIfEmpty();

            if (currentSlide.TextAreas.Last().Transform.SizeY > 0)
            {
                AddTextArea();
                lastTextArea = currentSlide.TextAreas.Last();
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

            if (LinkStack.Count > 0)
            {
                Text.Link = LinkStack.Peek();
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

        public PPTXTextArea AddTextArea()
        {
            WantReturn = false;

            currentSlide.TextAreas.Add(new PPTXTextArea(Settings.Margin.Left,
                LastAddedItemTransform.PositionY + LastAddedItemTransform.SizeY + Settings.TextAreaMarginHeight,
                PageWidth,
                0));


            return currentSlide.TextAreas.Last();
        }

        public void EndTextArea()
        {
            var lastTextArea = AddTextAreaIfEmpty();
            if (lastTextArea.Transform.SizeY > 0) return;

            var lastTextAreaSize = 0.0f;

            foreach (var _text in lastTextArea.Texts)
            {
                float maxFontHeight = 0;

                foreach (var _textRun in _text.Texts)
                {
                    maxFontHeight = Math.Max(maxFontHeight, FontHeght(_textRun.Font) * 1.2f);
                }

                lastTextAreaSize += maxFontHeight;
            }

            lastTextArea.Transform.SizeY = lastTextAreaSize;
            LastAddedItemTransform = lastTextArea.Transform;

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
                AddTextArea();
            }

            return currentSlide.TextAreas.Last();
        }

        public void PushHyperLink(PPTXLink Link)
        {
            LinkStack.Push(Link);
        }

        public void PopHyperLink()
        {
            LinkStack.Pop();
        }

        public void WriteImage(PPTXImage Image)
        {
            WantReturn = false;

            Image.Transform.PositionY = LastAddedItemTransform.PositionY + LastAddedItemTransform.SizeY + Settings.TextAreaMarginHeight;

            currentSlide.Images.Add(Image);

            LastAddedItemTransform = Image.Transform;
        }
    }
}
