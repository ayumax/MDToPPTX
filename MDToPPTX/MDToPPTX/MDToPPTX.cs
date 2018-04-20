using System;



namespace MDToPPTX
{
    public class MDToPPTX
    {
        public void Run(string PPTXFilePath)
        {
            using (PPTX.PPTXDocument document = new PPTX.PPTXDocument(PPTXFilePath))
            {
                var slide1 = new PPTX.PPTXSlide() { Body = "てすとぼでぃーーーーー" };
                document.AddSlide(slide1);
            }               
        }   
    }
}
