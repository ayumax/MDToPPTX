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


        public Stack<PPTXFont> FontStack { get; private set;} = new Stack<PPTXFont>();
        public Stack<PPTXLink> LinkStack { get; private set; } = new Stack<PPTXLink>();

        public PPTXTransform LastAddedItemTransform = new PPTXTransform();

        private SlideTextManager TextManager = new SlideTextManager();
        private SlideTableManager TableManager = new SlideTableManager();

        public PPTXTransform NewTransform => new PPTXTransform(Settings.Margin.Left,
                LastAddedItemTransform.PositionY + LastAddedItemTransform.SizeY,
                PageWidth,
                0);

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

            FontStack.Clear();
            LastAddedItemTransform = new PPTXTransform();

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

        public PPTXTextArea AddTextArea()
        {
            return TextManager.AddTextArea();
        }

        public void EndTextArea()
        {
            TextManager.EndTextArea();
        }

        public void PushFont(PPTXFont Font)
        {
            FontStack.Push(Font);
        }

        public void PopFont()
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
            Image.Transform = NewTransform;

            using (System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(Image.ImageFilePath))
            {
                Image.Transform.SizeX = bitmap.Width / 1000.0f;
                Image.Transform.SizeY = bitmap.Height / 1000.0f;
            }

            currentSlide.Images.Add(Image);

            LastAddedItemTransform = Image.Transform;
        }

        public void AddTable(PPTXTable Table)
        {
            PushFont(Settings.ListItemFont);

            currentSlide.Tables.Add(Table);
            TableManager.AddTable(Table);
        }

        public void AddTableEnd()
        {
            TableManager.AddTableEnd();

            PopFont();
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
    }
}
