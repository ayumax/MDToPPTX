using Markdig.Syntax;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    /// <summary>
    /// A PPTX renderer for a <see cref="ThematicBreakBlock"/>.
    /// </summary>
    public class ThematicBreakRenderer : PPTXObjectRenderer<ThematicBreakBlock>
    {
        protected override void Write(PPTXRenderer renderer, ThematicBreakBlock obj)
        {
            renderer.InsertNewPage();
        }
    }
}