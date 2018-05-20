using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for a <see cref="HtmlEntityInline"/>.
    /// </summary>
    public class PPTXHtmlEntityInlineRenderer : PPTXObjectRenderer<HtmlEntityInline>
    {
        protected override void Write(PPTXRenderer renderer, HtmlEntityInline obj)
        {
            renderer.Write(obj.Original);
        }
    }
}