using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Runtime.Serialization;

namespace MDToPPTX.PPTX
{
    public class PPTXColor
    {
        public int R { get; set; }
        public int G { get; set; }
        public int B { get; set; }
        public int A { get; set; }

        public PPTXColor()
        {
            R = 0x00;
            G = 0x00;
            B = 0x00;
            A = 0xFF;
        }

        public PPTXColor(int Red, int Green, int Blue, int Alpha = 255)
        {
            R = Red;
            G = Green;
            B = Blue;
            A = Alpha;
        }

        public PPTXColor(Color Color)
        {
            R = Color.R;
            G = Color.G;
            B = Color.B;
            A = Color.A;
        }

        [IgnoreDataMember]
        public bool IsTransparent => A == 0;

        [IgnoreDataMember]
        public Color Color => Color.FromArgb(A, R, G, B);
    }
}
