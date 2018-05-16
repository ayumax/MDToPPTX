using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public class PPTXTransform
    {
        public bool AutoLayout { get; set; } = true;
        /// <summary>
        /// 横位置(cm)
        /// </summary>
        public float PositionX { get; set; }
        /// <summary>
        /// 縦位置(cm)
        /// </summary>
        public float PositionY { get; set; }
        /// <summary>
        /// 横幅(cm)
        /// </summary>
        public float SizeX { get; set; }
        /// <summary>
        /// 縦幅(cm)
        /// </summary>
        public float SizeY { get; set; }

        public PPTXTransform()
        {

        }

        public PPTXTransform(float PositionX, float PositionY, float SizeX, float SizeY)
        {
            this.AutoLayout = false;
            this.PositionX = PositionX;
            this.PositionY = PositionY;
            this.SizeX = SizeX;
            this.SizeY = SizeY;
        }
    }
}
