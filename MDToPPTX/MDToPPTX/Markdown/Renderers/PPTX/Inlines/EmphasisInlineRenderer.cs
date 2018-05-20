using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for an <see cref="EmphasisInline"/>.
    /// </summary>
    public class EmphasisInlineRenderer : PPTXObjectRenderer<EmphasisInline>
    {
        protected override void Write(PPTXRenderer renderer, EmphasisInline obj)
        {
            var emphasisText = new string(obj.DelimiterChar, obj.IsDouble ? 2 : 1);
            renderer.Write(emphasisText);
            renderer.WriteChildren(obj);
            renderer.Write(emphasisText);
        }
    }
}