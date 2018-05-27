using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for a <see cref="LineBreakInline"/>.
    /// </summary>
    public class LineBreakInlineRenderer : PPTXObjectRenderer<LineBreakInline>
    {
        protected override void Write(PPTXRenderer renderer, LineBreakInline obj)
        {
            renderer.WriteReturn();
        }
    }
}