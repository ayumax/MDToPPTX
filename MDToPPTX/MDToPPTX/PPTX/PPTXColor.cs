using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace MDToPPTX.PPTX
{
    public class PPTXColor
    {
        public Color Color { get; set; }

        public PPTXColor()
        {

        }

        public PPTXColor(int Red, int Green, int Blue, int Alpha = 255)
        {
            Color = Color.FromArgb(Alpha, Red, Green, Blue);
        }

        public PPTXColor(Color Color)
        {
            this.Color = Color;
        }

        public bool IsTransparent => Color == Color.Transparent || Color.A == 0;
    }
}
