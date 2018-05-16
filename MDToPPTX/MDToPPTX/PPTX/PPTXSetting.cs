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

    public enum EPPTXSlideLayoutType
    {
        /// <summary>
        /// タイトル スライド
        /// </summary>
        TitleSlide,
        /// <summary>
        /// タイトルとコンテンツ
        /// </summary>
        TitleAndContents,
        /// <summary>
        /// セクション見出し
        /// </summary>
        SectionHeading,
        /// <summary>
        /// 2 つのコンテンツ
        /// </summary>
        TwoContents,
        /// <summary>
        /// 比較
        /// </summary>
        Comparison,
        /// <summary>
        /// タイトルのみ
        /// </summary>
        TitleOnly,
        /// <summary>
        /// 白紙
        /// </summary>
        BlankSheet,
        /// <summary>
        /// タイトル付きのコンテンツ
        /// </summary>
        ContentWithTitle,
        /// <summary>
        /// タイトル付きの図
        /// </summary>
        DiagramWithTitle,
        /// <summary>
        /// タイトルと縦書きテキスト
        /// </summary>
        TitleAndVerticalText,
        /// <summary>
        /// 縦書きタイトルと\n縦書きテキスト
        /// </summary>
        VerticalTitleAndVerticalText
    }
    public static partial class EPPTXSlideLayoutTypeExtend
    {
        public static string GetLayoutName(this EPPTXSlideLayoutType slideType)
        {
            string ret = "";
            switch (slideType)
            {
                case EPPTXSlideLayoutType.TitleSlide:
                    ret = "タイトル スライド";
                    break;
                case EPPTXSlideLayoutType.TitleAndContents:
                    ret = "タイトルとコンテンツ";
                    break;
                case EPPTXSlideLayoutType.SectionHeading:
                    ret = "セクション見出し";
                    break;
                case EPPTXSlideLayoutType.TwoContents:
                    ret = "2 つのコンテンツ";
                    break;
                case EPPTXSlideLayoutType.Comparison:
                    ret = "比較";
                    break;
                case EPPTXSlideLayoutType.TitleOnly:
                    ret = "タイトルのみ";
                    break;
                case EPPTXSlideLayoutType.BlankSheet:
                    ret = "白紙";
                    break;
                case EPPTXSlideLayoutType.ContentWithTitle:
                    ret = "タイトル付きのコンテンツ";
                    break;
                case EPPTXSlideLayoutType.DiagramWithTitle:
                    ret = "タイトル付きの図";
                    break;
                case EPPTXSlideLayoutType.TitleAndVerticalText:
                    ret = "タイトルと縦書きテキスト";
                    break;
                case EPPTXSlideLayoutType.VerticalTitleAndVerticalText:
                    ret = "縦書きタイトルと\n縦書きテキスト";
                    break;
            }
            return ret;
        }
    }

    public class PPTXSlideLayout
    {
        public EPPTXSlideLayoutType SlideType { get; set; } = EPPTXSlideLayoutType.BlankSheet;
        public string ID { get; set; }
        public string Name => SlideType.GetLayoutName();

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
    }

    public class PPTXSetting
    {
        public EPPTXSlideSizeValues SlideSize { get; set; } = EPPTXSlideSizeValues.Screen4x3;
        public Dictionary<EPPTXSlideLayoutType, PPTXSlideLayout> SlideLayouts { get; set; }

        public string Title { get; set; } = "無題";
        public string SubTitle { get; set; } = "-";

        public float SlideWidth { get; set; } = 25;
        public PPTXSlideMargin Margin { get; set; } = new PPTXSlideMargin();


        public PPTXFont Header1font { get; set; } = new PPTXFont() { FontSize = 32 };
        public PPTXFont Header2font { get; set; } = new PPTXFont() { FontSize = 28 };
        public PPTXFont NormalFont { get; set; } = new PPTXFont() { FontSize = 24 };
        public PPTXFont CodeFont { get; set; } = new PPTXFont() { FontSize = 12 };

        public PPTXSetting()
        {
            SlideLayouts = new Dictionary<EPPTXSlideLayoutType, PPTXSlideLayout>();
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.TitleSlide, ID = "rId1" });
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.TitleAndContents, ID = "rId2" });
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.SectionHeading, ID = "rId3" });
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.TwoContents, ID = "rId4" });
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.Comparison, ID = "rId5" });
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.TitleOnly, ID = "rId6" });
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.BlankSheet, ID = "rId7" });
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.ContentWithTitle, ID = "rId8" });
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.DiagramWithTitle, ID = "rId9" });
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.TitleAndVerticalText, ID = "rId10" });
            AddSlideLayout(new PPTXSlideLayout() { SlideType = EPPTXSlideLayoutType.VerticalTitleAndVerticalText, ID = "rId11" });
        }

        private void AddSlideLayout(PPTXSlideLayout Layout)
        {
            SlideLayouts.Add(Layout.SlideType, Layout);
        }
    }
}
