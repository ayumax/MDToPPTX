using System;
using System.Collections.Generic;
using System.IO;
using MDToPPTX.PPTX;
using MDToPPTX.Markdown;
using MDToPPTX.Markdown.SyntaxWriter;

namespace MDToPPTX
{
    public class MDToPPTX
    {
        private Dictionary<Type, SyntaxWriterBase> SyntaxWriter;

        public MDToPPTX()
        {
            SyntaxWriter = new Dictionary<Type, SyntaxWriterBase>();
            SyntaxWriter.Add(typeof(Markdig.Syntax.HeadingBlock), new HeadingBlockWriter());
            SyntaxWriter.Add(typeof(Markdig.Syntax.ParagraphBlock), new ParagraphBlockWriter());
        }

        public void Run(string MarkdownFilePath)
        {
            var markdownText = "";
            using (StreamReader sr = new StreamReader(MarkdownFilePath))
            {
                markdownText = sr.ReadToEnd();
            }

            var parsedMarkdown = Markdig.Markdown.Parse(markdownText);

            var settings = new PPTXSetting()
            {
                SlideSize = EPPTXSlideSizeValues.Screen4x3,
                Title = "パワポサンプル",
                SubTitle = "2018/5/3 ayumax"
            };

            using (PPTXDocument document = new PPTXDocument(MarkdownFilePath.ToLower().Replace(".md", ".pptx"), settings))
            {
                var slide = new SlideManager(document, settings);

                foreach (var _markdownItem in parsedMarkdown)
                {
                    var _markdownType = _markdownItem.GetType();
                    if (SyntaxWriter.ContainsKey(_markdownType))
                    {
                        SyntaxWriter[_markdownType].Write(_markdownItem, slide);
                    }
                }
            }

        }   

       
    }
}
