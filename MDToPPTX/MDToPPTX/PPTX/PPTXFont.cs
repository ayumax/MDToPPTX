using System;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public class PPTXFont
    {
        public string FontFamily { get; set; } = "メイリオ";
        public float FontSize { get; set; } = 28;
        public PPTXColor ForegroundColor { get; set; } = new PPTXColor(System.Drawing.Color.Black);
        public bool Bold { get; set; } = false;
        public bool Italic { get; set; } = false;
        public bool UnderLine { get; set; } = false;
        public bool Strike { get; set; } = false;
        public EPPTXHAlign HAlign { get; set; } = EPPTXHAlign.Left;

        public PPTXFont Clone()
        {
            var newObj = new PPTXFont();

            PropertyInfo[] infoArray = GetType().GetProperties();
            foreach (PropertyInfo info in infoArray)
            {
                info.SetValue(newObj, info.GetValue(this));
            }

            return newObj;
        }
    }
}
