using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for an <see cref="AutolinkInline"/>.
    /// </summary>
    public class AutolinkInlineRenderer : PPTXObjectRenderer<AutolinkInline>
    {
        protected override void Write(PPTXRenderer renderer, AutolinkInline obj)
        {
            renderer.Write(obj.Url);
        }
    }
}