using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public class PPTXText
    {
        /// <summary>
        /// 追加されるテキスト
        /// </summary>
        public string Text { get; set; } = "";

        /// <summary>
        /// テキストの横位置(cm)
        /// </summary>
        public int PositionX { get; set; }
        /// <summary>
        /// テキストの縦位置(cm)
        /// </summary>
        public int PositionY { get; set; }
        /// <summary>
        /// テキストの横幅(cm)
        /// </summary>
        public int SizeX { get; set; }
        /// <summary>
        /// テキストの縦幅(cm)
        /// </summary>
        public int SizeY { get; set; }

        public PPTXText()
        {

        }

        public PPTXText(string Text)
        {
            this.Text = Text;
        }
    }
}
