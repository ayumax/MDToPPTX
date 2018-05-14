using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MDToPPTX;
using MDToPPTX.PPTX;

namespace MDToPPTXSample
{
    class Program
    {
        static void Main(string[] args)
        {
            MDToPPTX.MDToPPTX pptxConverter = new MDToPPTX.MDToPPTX();

            string filepath = @"C:\Users\ayuma\Desktop\sample3.pptx";

            pptxConverter.Run(filepath);
        }

        private void OutputPPTXDirect(string PPTXFilePath)
        {
            var settings = new PPTXSetting()
            {
                SlideSize = EPPTXSlideSizeValues.Screen4x3,
                Title = "パワポサンプル",
                SubTitle = "2018/5/3 ayumax"
            };

            using (PPTXDocument document = new PPTXDocument(PPTXFilePath, settings))
            {
                document.Slides = new List<PPTXSlide>()
                {
                    new PPTXSlide()
                    {
                        SlideLayout = settings.SlideLayouts[EPPTXSlideLayoutType.TitleAndContents],
                        Title = new PPTXTextArea("コンテンツ１ページ目"),
                        TextAreas = new List<PPTXTextArea>()
                        {
                            new PPTXTextArea("本文です。\n\\nをいれると改行もされます")
                        }
                    },
                    new PPTXSlide()
                    {
                        SlideLayout = settings.SlideLayouts[EPPTXSlideLayoutType.TitleOnly],
                        Title = new PPTXTextArea("コンテンツ２ページ目"),
                        TextAreas = new List<PPTXTextArea>()
                        {
                            new PPTXTextArea("パワーポイント2枚目のテキスト１", 1, 5, 20, 2),
                            new PPTXTextArea(1, 7, 20, 7)
                            {
                                Texts = new List<PPTXText>()
                                {
                                    new PPTXText("2枚目1行目", PPTXBullet.Circle),
                                    new PPTXText("2枚目2行目", PPTXBullet.Circle),
                                    new PPTXText("2枚目3行目", PPTXBullet.Rectangle),
                                    new PPTXText("2枚目4行目 箇条書き解除")
                                }
                            }
                        },
                        Images = new List<PPTXImage>()
                        {
                            new PPTXImage(@"C:\temp\sample.jpg", 1, 15, 5, 3),
                            new PPTXImage(@"C:\temp\sample.jpg", 7, 15, 5, 3)
                        }
                    }
                };
            }
        }
    }
}
