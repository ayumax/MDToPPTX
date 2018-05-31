using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MDToPPTX.PPTX;

namespace MDToPPTX.Markdown
{
    class SlideTextManager
    {
        private SlideManager SlideManager;
        private bool WantReturn = false;

        public void Init(SlideManager SlideManager)
        {
            this.SlideManager = SlideManager;

            WantReturn = false;

        }

        public void Write(PPTXTextRun Text)
        {
            var lastTextArea = AddTextAreaIfEmpty();

            if (SlideManager.currentSlide.TextAreas.Last().Transform.SizeY > 0)
            {
                AddTextArea();
                lastTextArea = SlideManager.currentSlide.TextAreas.Last();
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

            if (SlideManager.FontStack.Count > 0)
            {
                Text.Font = SlideManager.FontStack.Peek();
            }

            if (SlideManager.LinkStack.Count > 0)
            {
                Text.Link = SlideManager.LinkStack.Peek();
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

            var newTextArea = new PPTXTextArea(SlideManager.NewTransform);
            SlideManager.currentSlide.TextAreas.Add(newTextArea);

            return newTextArea;
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
                    maxFontHeight = Math.Max(maxFontHeight, SlideManager.FontHeght(_textRun.Font) * 1.2f);
                }

                lastTextAreaSize += maxFontHeight;
            }

            lastTextArea.Transform.SizeY = lastTextAreaSize;
            SlideManager.LastAddedItemTransform = lastTextArea.Transform;

            WantReturn = false;
        }

        private PPTXTextArea AddTextAreaIfEmpty()
        {
            if (SlideManager.currentSlide.TextAreas.Count == 0)
            {
                AddTextArea();
            }

            return SlideManager.currentSlide.TextAreas.Last();
        }
    }
}
