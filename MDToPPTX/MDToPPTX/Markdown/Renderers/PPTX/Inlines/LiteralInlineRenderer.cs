using Markdig.Helpers;
using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for a <see cref="LiteralInline"/>.
    /// </summary>
    public class LiteralInlineRenderer : PPTXObjectRenderer<LiteralInline>
    {
        protected override void Write(PPTXRenderer renderer, LiteralInline obj)
        {
            renderer.Write(ref obj.Content);
        }
    }
}