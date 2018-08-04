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
        public float PageHeight => Settings.SlideHeight - (Settings.Margin.Top + Settings.Margin.Bottom);

        public PPTXBlockSetting BlockSetting { get; private set; } = new PPTXBlockSetting();
        public Stack<PPTXFont> FontStack { get; private set; } = new Stack<PPTXFont>();
        public Stack<PPTXLink> LinkStack { get; private set; } = new Stack<PPTXLink>();

        public PPTXTransform LastAddedItemTransform = new PPTXTransform();

        private SlideTextManager TextManager = new SlideTextManager();
        private SlideTableManager TableManager = new SlideTableManager();

        public SlideManager(PPTXDocument document, PPTXSetting Settings)
        {
            this.document = document;
            this.Settings = Settings;

            CreateNewSlide();
        }

        public PPTXSlide CreateNewSlide()
        {
            if (currentSlide != null)
            {
                EndSheet(currentSlide);
            }
            currentSlide = new PPTXSlide() { SlideLayout = EPPTXSlideLayoutType.BlankSheet };
            document.Slides.Add(currentSlide);

            FontStack.Clear();
            LastAddedItemTransform = new PPTXTransform();
            LastAddedItemTransform.SizeY = Settings.Margin.Top;


            TextManager.Init(this);
            TableManager.Init(this);

            return currentSlide;
        }

        public void Write(PPTXTextRun Text)
        {
            if (TableManager.IsReadyCell)
            {
                TableManager.Write(Text);
            }
            else
            {
                TextManager.Write(Text);
            }
        }

        public void AddTextRow(PPTXText Text)
        {
            TextManager.AddTextRow(Text);
        }

        public void WriteReturn()
        {
            TextManager.WriteReturn();
        }

        public void AddTextArea()
        {
            TextManager.AddTextArea();
        }

        public void EndTextArea()
        {
            TextManager.EndTextArea();
        }

        public void PushBlockSetting(PPTXBlockSetting BlockSetting)
        {
            this.BlockSetting = BlockSetting;
            FontStack.Push(BlockSetting.Font);
        }

        public void PopBlockSetting()
        {
            this.BlockSetting = null;
            FontStack.Pop();
        }

        public void PushInlineSetting(PPTXInlineSetting InlieSetting)
        {
            FontStack.Push(InlieSetting.Font);
        }

        public void PopInline()
        {
            FontStack.Pop();
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
            Image.Transform = NewTransform();

            using (System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(Image.ImageFilePath))
            {
                // 1mm = 3.779528px(96dpi)
                Image.Transform.SizeX = bitmap.Width / 3.779528f / 10.0f;
                Image.Transform.SizeY = bitmap.Height / 3.779528f / 10.0f;
            }

            currentSlide.Images.Add(Image);

            SetContentTransform(Image.Transform);
        }

        public void AddTable(PPTXTable Table)
        {
            currentSlide.Tables.Add(Table);
            TableManager.AddTable(Table);
        }

        public void AddTableEnd()
        {
            TableManager.AddTableEnd();
        }

        public void AddTableRow()
        {
            TableManager.AddTableRow();
        }

        public void NextTableCell()
        {
            TableManager.NextTableCell();
        }

        public void EndTableRow()
        {
            TableManager.EndTableRow();
        }

        public PPTXTransform NewTransform()
        {
            var margin = BlockSetting?.Margin ?? new PPTXMargin();

            return new PPTXTransform(Settings.Margin.Left + margin.Left,
               LastAddedItemTransform.PositionY + LastAddedItemTransform.SizeY + margin.Top,
               PageWidth - margin.Right,
               0);
        }

        public PPTXFont CurrentFont => FontStack.Count > 0 ? FontStack.Peek() : Settings.Normal.Font;

        public void SetContentTransform(PPTXTransform transform)
        {
            LastAddedItemTransform = transform;
        }

        public void EndSheet(PPTXSlide CurrentSheet = null)
        {
            var _CurrentSheet = CurrentSheet ?? currentSlide;

            if (Settings.StackLayoutRule == EPPTXSideStackLayoutRule.Center)
            {
                var slideCenterY = PageHeight / 2.0f;
                var contentCentetY = ((LastAddedItemTransform.PositionY + LastAddedItemTransform.SizeY) - Settings.Margin.Top) / 2.0f + Settings.Margin.Top;
                var diffY = slideCenterY - contentCentetY;

                _CurrentSheet.TextAreas.ForEach(_textArea => _textArea.Transform.PositionY += diffY);
                _CurrentSheet.Tables.ForEach(_table => _table.Transform.PositionY += diffY);

                var slideCenterX = PageWidth / 2.0f;
                _CurrentSheet.Images.ForEach(_image => {
                    var diffX = slideCenterX - (_image.Transform.PositionX + _image.Transform.SizeX / 2);
                    _image.Transform.PositionX += diffX;
                    _image.Transform.PositionY += diffY;
                });
            }
        }
    }
}
