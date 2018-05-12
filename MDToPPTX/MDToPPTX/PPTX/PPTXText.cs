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
        /// テキストの位置
        /// </summary>
        public PPTXTransform Transform { get; set; } = new PPTXTransform();

        public PPTXText()
        {

        }

        public PPTXText(string Text)
        {
            this.Text = Text;
        }
    }
}
