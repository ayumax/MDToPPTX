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

    public class PPTXSlideMargin
    {
        public float Left { get; set; } = 1;
        public float Top { get; set; } = 1;
        public float Right { get; set; } = 1;
        public float Bottom { get; set; } = 1;
    }

    public class PPTXFont
    {
        public string FontFamily { get; set; } = "メイリオ";
        public float FontSize { get; set; } = 28;
        public PPTXColor ForegroundColor { get; set; } = new PPTXColor(System.Drawing.Color.Black);
    }

    public class PPTXSetting
    {
        public EPPTXSlideSizeValues SlideSize { get; set; } = EPPTXSlideSizeValues.Screen4x3;
       
        public string Title { get; set; } = "無題";
        public string SubTitle { get; set; } = "-";

        public float SlideWidth { get; set; } = 25;
        public PPTXSlideMargin Margin { get; set; } = new PPTXSlideMargin();


        public PPTXFont Header1font { get; set; } = new PPTXFont() { FontSize = 32, ForegroundColor = new PPTXColor(0, 0, 0) };
        public PPTXFont Header2font { get; set; } = new PPTXFont() { FontSize = 28 };
        public PPTXFont NormalFont { get; set; } = new PPTXFont() { FontSize = 24 };
        public PPTXFont CodeFont { get; set; } = new PPTXFont() { FontSize = 12 };
        public PPTXFont InlineCodeFont { get; set; } = new PPTXFont() { FontSize = 24, ForegroundColor = new PPTXColor(204, 51, 0) };
        public PPTXFont ListItemFont { get; set; } = new PPTXFont() { FontSize = 18 };

        public PPTXSetting()
        {
        }

    }
}
