using System;

namespace MDToPPTX
{
    public class MDToPPTX
    {
        public void Run(string PPTXFilePath)
        {
            var settings = new PPTX.PPTXSetting()
            {
                SlideSize = PPTX.EPPTXSlideSizeValues.Screen4x3,
                Title = "たいとるABCDEFG",
                SubTitle = "2018/5/3 ayumax"
            };

            using (PPTX.PPTXDocument document = new PPTX.PPTXDocument(PPTXFilePath, settings))
            {
                document.Slides = new System.Collections.Generic.List<PPTX.PPTXSlide>()
                {
                    new PPTX.PPTXSlide()
                    {
                        SlideLayout = settings.SlideLayouts[PPTX.EPPTXSlideLayoutType.TitleAndContents],
                        Title = new PPTX.PPTXText("コンテンツ１つめ"),
                        Bodys = new System.Collections.Generic.List<PPTX.PPTXText>()
                        {
                            new PPTX.PPTXText(){ Text = "てすとぼでぃーーーーー", PositionX = 0, PositionY = 0, SizeX = 10, SizeY = 2 },
                            new PPTX.PPTXText(){ Text = "てすとぼでぃーーーーー2", PositionX = 0, PositionY = 2, SizeX = 10, SizeY = 2 }
                        }
                    },
                    new PPTX.PPTXSlide()
                    {
                        SlideLayout = settings.SlideLayouts[PPTX.EPPTXSlideLayoutType.TwoContents],
                        Title = new PPTX.PPTXText("コンテンツ２つめ"),
                        Bodys = new System.Collections.Generic.List<PPTX.PPTXText>()
                        {
                            new PPTX.PPTXText(){ Text = "パワーポイント2枚目のテキスト１" },
                            new PPTX.PPTXText(){ Text = "テキスト２\r\n２行目" }
                        }
                    }
                };
            }               
        }   
    }
}
