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

        /// <summary>
        /// Convert Markdown text
        /// </summary>
        /// <param name="MarkdownFilePath"></param>
        /// <param name="options"></param>
        public void RunFromMDFile(string MarkdownFilePath, string ExportPath = null, PPTXSetting options = null)
        {
            var markdownText = "";
            using (StreamReader sr = new StreamReader(MarkdownFilePath))
            {
                markdownText = sr.ReadToEnd();
            }

            RunFromMDText(markdownText, ExportPath ?? MarkdownFilePath.ToLower().Replace(".md", ".pptx"), options);
        }

        /// <summary>
        /// Convert Markdown text
        /// </summary>
        /// <param name="MarkdownText">Markdown text</param>
        /// <param name="ExportPath">pptx file path</param>
        /// <param name="options">Option setting</param>
        public void RunFromMDText(string MarkdownText, string ExportPath, PPTXSetting options = null) => ToPPTX(MarkdownText, ExportPath, options);

        protected static void ToPPTX(string markdown, string pptxFilePath, PPTXSetting options = null, MarkdownPipeline pipeline = null)
        {
            var pptx = ToPPTxDocument(markdown, options, pipeline);

            pptx.SaveAs(pptxFilePath, options);
        }

        /// <summary>
        /// Make pptx document
        /// </summary>
        /// <param name="MarkdownFilePath"></param>
        /// <param name="options"></param>
        public PPTXDocument MakePPtxDocumentFromMDFile(string MarkdownFilePath, PPTXSetting options = null)
        {
            var markdownText = "";
            using (StreamReader sr = new StreamReader(MarkdownFilePath))
            {
                markdownText = sr.ReadToEnd();
            }

            return MakePPtxDocumentFromMDText(markdownText, options);
        }

        /// <summary>
        /// Convert Markdown text
        /// </summary>
        /// <param name="MarkdownText">Markdown text</param>
        /// <param name="ExportPath">pptx file path</param>
        /// <param name="options">Option setting</param>
        public PPTXDocument MakePPtxDocumentFromMDText(string MarkdownText, PPTXSetting options = null) => ToPPTxDocument(MarkdownText, options);

        protected static PPTXDocument ToPPTxDocument(string markdown, PPTXSetting options = null, MarkdownPipeline pipeline = null)
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

            var pptx = new PPTXDocument();

            var slide = new SlideManager(pptx, options);

            var renderer = new PPTXRenderer(slide, options);
            pipeline.Setup(renderer);

            renderer.Render(document);

            slide.EndSheet();

            return pptx;
        }
    }
}
