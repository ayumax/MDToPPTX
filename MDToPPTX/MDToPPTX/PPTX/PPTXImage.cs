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

        public PPTXImage(string ImageFilePath)
        {
            this.ImageFilePath = ImageFilePath;
        }

        public PPTXImage(string ImageFilePath, float PositionX, float PositionY, float SizeX, float SizeY)
        {
            this.ImageFilePath = ImageFilePath;

            this.Transform = new PPTXTransform()
            {
                AutoLayout = false,
                PositionX = PositionX,
                PositionY = PositionY,
                SizeX = SizeX,
                SizeY = SizeY
            };
        }
    }
}
