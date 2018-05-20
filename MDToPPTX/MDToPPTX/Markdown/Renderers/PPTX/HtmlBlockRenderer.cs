using Markdig.Syntax;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    public class HtmlBlockRenderer : PPTXObjectRenderer<HtmlBlock>
    {
        protected override void Write(PPTXRenderer renderer, HtmlBlock obj)
        {
            renderer.WriteLeafRawLines(obj, true, false);
        }
    }
}