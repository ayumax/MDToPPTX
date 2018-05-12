using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public class PPTXImage
    {
        /// <summary>
        /// 画像ファイルパス
        /// </summary>
        public string ImageFilePath { get; set; } = "";

        /// <summary>
        /// イメージの位置
        /// </summary>
        public PPTXTransform Transform { get; set; } = new PPTXTransform();
    }
}
