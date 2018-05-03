using System;



namespace MDToPPTX
{
    public class MDToPPTX
    {
        public void Run(string PPTXFilePath)
        {
            using (PPTX.PPTXDocument document = new PPTX.PPTXDocument(PPTXFilePath, "たいとるABCDEFG", "2018/5/3 ayumax"))
            {
                document.Slides = new System.Collections.Generic.List<PPTX.PPTXSlide>()
                {
                    new PPTX.PPTXSlide()
                    {
                        Bodys = new System.Collections.Generic.List<PPTX.PPTXText>()
                        {
                            new PPTX.PPTXText(){ Text = "てすとぼでぃーーーーー", PositionX = 0, PositionY = 0, SizeX = 10, SizeY = 2 },
                            new PPTX.PPTXText(){ Text = "てすとぼでぃーーーーー2", PositionX = 0, PositionY = 2, SizeX = 10, SizeY = 2 }
                        }
                    },
                    new PPTX.PPTXSlide()
                    {
                        Bodys = new System.Collections.Generic.List<PPTX.PPTXText>()
                        {
                            new PPTX.PPTXText(){ Text = "パワーポイント2枚目のテキスト１", PositionX = 0, PositionY = 0, SizeX = 10, SizeY = 2 },
                            new PPTX.PPTXText(){ Text = "テキスト２\r\n２行目", PositionX = 2, PositionY = 2, SizeX = 10, SizeY = 2 }
                        }
                    }
                };
            }               
        }   
    }
}
