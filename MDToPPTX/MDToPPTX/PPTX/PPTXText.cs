using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    /// <summary>
    /// テキスト最小単位
    /// </summary>
    public class PPTXTextRun
    {
        /// <summary>
        /// 追加されるテキスト
        /// </summary>
        public string Text { get; set; } = "";
        public PPTXFont Font { get; set; } = new PPTXFont();
        public PPTXColor ForegroundColor { get; set; } = new PPTXColor() { Color = System.Drawing.Color.Black };
    }

    /// <summary>
    /// テキストエリア内の1行
    /// </summary>
    public class PPTXText
    {
        /// <summary>
        /// Texts
        /// </summary>
        public List<PPTXTextRun> Texts { get; set; } = new List<PPTXTextRun>();

        /// <summary>
        /// 箇条書き設定
        /// </summary>
        public PPTXBullet Bullet { get; set; } = PPTXBullet.None;

        public PPTXText()
        {

        }

        public PPTXText(PPTXBullet Bullet)
        {
            this.Bullet = Bullet;
        }

        public PPTXText(PPTXBullet Bullet, PPTXTextRun Text)
        {
            this.Bullet = Bullet;

            Texts.Add(Text);
        }
    }

    /// <summary>
    /// テキストエリア
    /// </summary>
    public class PPTXTextArea
    {
        /// <summary>
        /// テキストの位置
        /// </summary>
        public PPTXTransform Transform { get; set; } = new PPTXTransform();

        /// <summary>
        /// 箇条書き設定
        /// </summary>
        public List<PPTXText> Texts { get; set; } = new List<PPTXText>();

        public PPTXColor BackgroundColor { get; set; } = new PPTXColor() { Color = System.Drawing.Color.Transparent };

        public PPTXTextArea()
        {

        }

        public PPTXTextArea(float PositionX, float PositionY, float SizeX, float SizeY)
        {
            this.Texts = new List<PPTXText>();

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
