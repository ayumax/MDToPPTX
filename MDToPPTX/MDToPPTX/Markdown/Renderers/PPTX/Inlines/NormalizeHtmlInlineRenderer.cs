using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for a <see cref="HtmlInline"/>.
    /// </summary>
    public class PPTXHtmlInlineRenderer : PPTXObjectRenderer<HtmlInline>
    {
        protected override void Write(PPTXRenderer renderer, HtmlInline obj)
        {
            renderer.Write(obj.Tag);
        }
    }
}