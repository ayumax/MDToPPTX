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
    public class MDToPPTX
    {
        public MDToPPTX()
        {
        }

        public void Run(string MarkdownFilePath, string PPTitle = "", string PPSubTitle = "")
        {
            var markdownText = "";
            using (StreamReader sr = new StreamReader(MarkdownFilePath))
            {
                markdownText = sr.ReadToEnd();
            }

            var settings = new PPTXSetting()
            {
                SlideSize = EPPTXSlideSizeValues.Screen4x3,
                Title = string.IsNullOrWhiteSpace(PPTitle) ? Path.GetFileNameWithoutExtension(MarkdownFilePath) : PPTitle,
                SubTitle = PPSubTitle
            };

            ToPPTX(markdownText, MarkdownFilePath.ToLower().Replace(".md", ".pptx"), settings);
        }

        public static MarkdownDocument ToPPTX(string markdown, string pptxFilePath, PPTXSetting options = null, MarkdownPipeline pipeline = null)
        {
            pipeline = pipeline ?? new MarkdownPipelineBuilder()
                .UsePipeTables()
                .UseEmphasisExtras()
                .Build();
            //pipeline = Markdig.Markdown.CheckForSelfPipeline(pipeline, markdown); 

            var document = Markdig.Markdown.Parse(markdown, pipeline);

            using (PPTXDocument pptx = new PPTXDocument(pptxFilePath, options))
            {
                var slide = new SlideManager(pptx, options);

                // We override the renderer with our own writer
                var renderer = new PPTXRenderer(slide, options);
                pipeline.Setup(renderer);

                renderer.Render(document);
            }

            return document;
        }
    }
}
