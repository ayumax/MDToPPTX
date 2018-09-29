using System;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public enum EPPTXSlideSizeValues
    {
        Screen4x3,
        Screen16x9,
    }

    public enum EPPTXSideStackLayoutRule
    {
        /// <summary>
        /// Layout from upper left to lower left
        /// </summary>
        Normal,
        /// <summary>
        /// Layout gathering in the middle
        /// </summary>
        Center,

        /// <summary>
        /// Make only the page with only one item the center
        /// </summary>
        CenterWithOnlyOneItem
    }

    public enum EPPTXHAlign
    {
        Left,
        Center,
        Right
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
       
        public string Title { get; set; } = "No title";
        public string SubTitle { get; set; } = "-";

        public float SlideWidth { get; set; } = 25.4f;
        public float SlideHeight { get; set; } = 19.05f;
        public PPTXMargin Margin { get; set; } = new PPTXMargin(1, 1, 1, 1);

        public EPPTXSideStackLayoutRule StackLayoutRule { get; set; } = EPPTXSideStackLayoutRule.Normal;


        public PPTXBlockSetting Header1 { get; set; } = new PPTXBlockSetting()
        {
            Font = new PPTXFont() { FontSize = 32, ForegroundColor = new PPTXColor(0, 0, 0), Bold = true, UnderLine = true, HAlign = EPPTXHAlign.Center }
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

        public static PPTXSetting Load(string LoadPath)
        {
            PPTXSetting retThis = null;

            using (var reader = new System.IO.StreamReader(LoadPath))
            {
                retThis = Utf8Json.JsonSerializer.Deserialize<PPTXSetting>(reader.BaseStream);
            }

            return retThis;
        }

        public void Save(string SavePath)
        {
            if (string.IsNullOrWhiteSpace(SavePath)) return;

            var serializedBuffer = Utf8Json.JsonSerializer.Serialize(this);

            using (var writer = new System.IO.FileStream(SavePath, System.IO.FileMode.Create))
            {
                writer.Write(serializedBuffer, 0, serializedBuffer.Length);
            }
        }

    }
}
