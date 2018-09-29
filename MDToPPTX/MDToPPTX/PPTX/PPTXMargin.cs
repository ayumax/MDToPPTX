using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
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
}
