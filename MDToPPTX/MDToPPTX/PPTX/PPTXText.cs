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
        /// 箇条書き設定
        /// </summary>
        public PPTXBullet Bullet { get; set; } = PPTXBullet.None;

        public int FontSize { get; set; } = 28;

        public string FontFamily { get; set; } = "メイリオ";

        public PPTXText()
        {

        }

        public PPTXText(string Text)
        {
            this.Text = Text;
            this.Bullet = PPTXBullet.None;
        }

        public PPTXText(string Text, PPTXBullet Bullet)
        {
            this.Text = Text;

            this.Bullet = Bullet;
        }
    }

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

        public PPTXTextArea()
        {

        }

        public PPTXTextArea(string Text)
        {
            this.Texts = new List<PPTXText>()
            {
                new PPTXText(Text)
            };

            this.Transform = new PPTXTransform();
        }

        public PPTXTextArea(string Text, int PositionX, int PositionY, int SizeX, int SizeY)
        {
            this.Texts = new List<PPTXText>()
            {
                new PPTXText(Text)
            };

            this.Transform = new PPTXTransform()
            {
                AutoLayout = false,
                PositionX = PositionX,
                PositionY = PositionY,
                SizeX = SizeX,
                SizeY = SizeY
            };
        }

        public PPTXTextArea(int PositionX, int PositionY, int SizeX, int SizeY)
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
