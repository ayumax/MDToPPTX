using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public enum EPPTXSlideSizeValues
    {
        Screen4x3,
        Screen16x9,
    }

    public class PPTXMargin
    {
        public float Left { get; set; } = 0;
        public float Top { get; set; } = 0;
        public float Right { get; set; } = 0;
        public float Bottom { get; set; } = 0;

        public PPTXMargin(float Left = 0, float Top = 0, float Right = 0, float Bottom = 0)
        {
            this.Left = Left;
            this.Top = Top;
            this.Right = Right;
            this.Bottom = Bottom;
        }
    }

    public class PPTXFont
    {
        public string FontFamily { get; set; } = "メイリオ";
        public float FontSize { get; set; } = 28;
        public PPTXColor ForegroundColor { get; set; } = new PPTXColor(System.Drawing.Color.Black);
    }

    public class PPTXInlineSetting
    {
        public PPTXFont Font { get; set; } = new PPTXFont();
    }

    public class PPTXBlockSetting
    {
        public PPTXMargin Margin { get; set; } = new PPTXMargin();
        public PPTXFont Font { get; set; } = new PPTXFont();
        public PPTXColor Background { get; set; } = new PPTXColor(System.Drawing.Color.Transparent);

    }

    public class PPTXSetting
    {
        public EPPTXSlideSizeValues SlideSize { get; set; } = EPPTXSlideSizeValues.Screen4x3;
       
        public string Title { get; set; } = "無題";
        public string SubTitle { get; set; } = "-";

        public float SlideWidth { get; set; } = 25;
        public PPTXMargin Margin { get; set; } = new PPTXMargin(1, 1, 1, 1);


        public PPTXBlockSetting Header1 { get; set; } = new PPTXBlockSetting()
        {
            Font = new PPTXFont() { FontSize = 32, ForegroundColor = new PPTXColor(0, 0, 0) }
        };

        public PPTXBlockSetting Header2 { get; set; } = new PPTXBlockSetting()
        {
            Font = new PPTXFont() { FontSize = 28, ForegroundColor = new PPTXColor(0, 0, 0) },
            Margin = new PPTXMargin(0.25f, 0.5f) 
        };

        public PPTXBlockSetting Normal { get; set; } = new PPTXBlockSetting()
        {
            Font = new PPTXFont() { FontSize = 24, ForegroundColor = new PPTXColor(0, 0, 0) },
            Margin = new PPTXMargin(0.5f, 0.5f) 
        };

        public PPTXBlockSetting Code { get; set; } = new PPTXBlockSetting()
        {
            Font = new PPTXFont() { FontSize = 12, ForegroundColor = new PPTXColor(0, 0, 0) },
            Margin = new PPTXMargin(0.5f, 0.5f, 0.5f),
            Background = new PPTXColor(245, 245, 245)
        };

        public PPTXBlockSetting List { get; set; } = new PPTXBlockSetting()
        {
            Font = new PPTXFont() { FontSize = 22, ForegroundColor = new PPTXColor(0, 0, 0) },
            Margin = new PPTXMargin(0.5f, 0.5f, 0.5f),
        };

        public PPTXBlockSetting QuoteBlock { get; set; } = new PPTXBlockSetting()
        {
            Font = new PPTXFont() { FontSize = 24, ForegroundColor = new PPTXColor(0, 0, 0) },
            Margin = new PPTXMargin(0.5f, 0.5f),
            Background = new PPTXColor(214, 220, 229)
        };

        public PPTXBlockSetting Table { get; set; } = new PPTXBlockSetting()
        {
            Font = new PPTXFont() { FontSize = 22, ForegroundColor = new PPTXColor(0, 0, 0) },
            Margin = new PPTXMargin(0.5f, 0.5f, 0.5f),
        };


        public PPTXInlineSetting InlineCode { get; set; } = new PPTXInlineSetting()
        {
            Font = new PPTXFont() { FontSize = 24, ForegroundColor = new PPTXColor(204, 51, 0) }
        };

        public PPTXSetting()
        {
        }

    }
}
