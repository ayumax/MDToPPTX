using System;
using System.Collections.Generic;
using System.IO;
using Markdig;
using Markdig.Renderers.Normalize;
using Markdig.Syntax;

using MDToPPTX.PPTX;
using MDToPPTX.Markdown;
using MDToPPTX.Markdown.Renderers.PPTX;

namespace MDToPPTX
{
    public class MD2PPTX
    {
        public MD2PPTX()
        {
        }

        public void Run(string MarkdownFilePath, PPTXSetting options = null)
        {
            var markdownText = "";
            using (StreamReader sr = new StreamReader(MarkdownFilePath))
            {
                markdownText = sr.ReadToEnd();
            }

            ToPPTX(markdownText, MarkdownFilePath.ToLower().Replace(".md", ".pptx"), options);
        }

        public static MarkdownDocument ToPPTX(string markdown, string pptxFilePath, PPTXSetting options = null, MarkdownPipeline pipeline = null)
        {
            options = options ?? new PPTXSetting()
            {
                SlideSize = EPPTXSlideSizeValues.Screen4x3
            };

            pipeline = pipeline ?? new MarkdownPipelineBuilder()
                .UsePipeTables()
                .UseEmphasisExtras()
                .Build();

            var document = Markdig.Markdown.Parse(markdown, pipeline);

            using (PPTXDocument pptx = new PPTXDocument(pptxFilePath, options))
            {
                var slide = new SlideManager(pptx, options);

                var renderer = new PPTXRenderer(slide, options);
                pipeline.Setup(renderer);

                renderer.Render(document);
            }

            return document;
        }
    }
}
