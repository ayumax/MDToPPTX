using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for a <see cref="DelimiterInline"/>.
    /// </summary>
    public class DelimiterInlineRenderer : PPTXObjectRenderer<DelimiterInline>
    {
        protected override void Write(PPTXRenderer renderer, DelimiterInline obj)
        {
            renderer.Write(obj.ToLiteral());
            renderer.WriteChildren(obj);
        }
    }
}