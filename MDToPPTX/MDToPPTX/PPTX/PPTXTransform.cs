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
        public int PositionX { get; set; }
        /// <summary>
        /// 縦位置(cm)
        /// </summary>
        public int PositionY { get; set; }
        /// <summary>
        /// 横幅(cm)
        /// </summary>
        public int SizeX { get; set; }
        /// <summary>
        /// 縦幅(cm)
        /// </summary>
        public int SizeY { get; set; }

        public PPTXTransform()
        {

        }
    }
}
