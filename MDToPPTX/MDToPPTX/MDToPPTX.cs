using System;
using System.Collections.Generic;
using MDToPPTX.PPTX;

namespace MDToPPTX
{
    public class MDToPPTX
    {
        public void Run(string PPTXFilePath)
        {
            var settings = new PPTXSetting()
            {
                SlideSize = EPPTXSlideSizeValues.Screen4x3,
                Title = "サンプルファイルタイトル",
                SubTitle = "2018/5/3 ayumax"
            };

            using (PPTXDocument document = new PPTXDocument(PPTXFilePath, settings))
            {
                document.Slides = new List<PPTXSlide>()
                {
                    new PPTXSlide()
                    {
                        SlideLayout = settings.SlideLayouts[EPPTXSlideLayoutType.TitleAndContents],
                        Title = new PPTXText("コンテンツ１ページ目"),
                        Texts = new List<PPTXText>()
                        {
                            new PPTXText(){ Text = "本文です。\nここに書いていく" }
                        },
                        Images = new List<PPTXImage>()
                        {
                            new PPTXImage(){ ImageFilePath = @"C:\Users\ayuma\Pictures\P7051318.JPG",
                                Transform = new PPTXTransform() { AutoLayout = false, PositionX = 1, PositionY = 7, SizeX = 10, SizeY = 6 }
                            }

                        }
                    },
                    new PPTXSlide()
                    {
                        SlideLayout = settings.SlideLayouts[EPPTXSlideLayoutType.TitleAndContents],
                        Texts = new List<PPTXText>()
                        {
                            new PPTXText(){ Text = "パワーポイント2枚目のテキスト１",
                                            Transform = new PPTXTransform() { AutoLayout = false, PositionX = 1, PositionY = 3, SizeX = 20, SizeY = 3 }},
                            new PPTXText(){ Text = "テキスト２\r\n２行目",
                            Transform = new PPTXTransform() { AutoLayout = false, PositionX = 1, PositionY = 6, SizeX = 10, SizeY = 6 } }
                        },
                        Images = new List<PPTXImage>()
                        {
                            new PPTXImage(){ ImageFilePath = @"C:\Users\ayuma\Pictures\ue1.PNG",
                            Transform = new PPTXTransform() { AutoLayout = false, PositionX = 7, PositionY = 4, SizeX = 10, SizeY = 6 }
                            }
                        }
                    }
                };
            }               
        }   
    }
}
