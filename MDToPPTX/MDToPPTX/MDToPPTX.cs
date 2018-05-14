using System;
using System.Collections.Generic;
using MDToPPTX.PPTX;

namespace MDToPPTX
{
    public class MDToPPTX
    {
        public void Run(string MarkdownFilePath)
        {
            var parsedMarkdown = Markdig.Markdown.Parse(
                "# sample markdown\n" + 
                "てすとです\n" + 
                "てすとなんです\n" + 
                "\n" + 
                "---\n" +
                "\n" +
                "箇条書き\n" +
                "あああ");

            var settings = new PPTXSetting()
            {
                SlideSize = EPPTXSlideSizeValues.Screen4x3,
                Title = "パワポサンプル",
                SubTitle = "2018/5/3 ayumax"
            };

            using (PPTXDocument document = new PPTXDocument(MarkdownFilePath.ToLower().Replace(".md", ".pptx"), settings))
            {
                document.Slides = new List<PPTXSlide>();
                var currentSlide = new PPTXSlide();

                foreach (var _markdownItem in parsedMarkdown)
                {
                    System.Diagnostics.Debug.WriteLine(_markdownItem);
                }

            }

           
        }   

       
    }
}
