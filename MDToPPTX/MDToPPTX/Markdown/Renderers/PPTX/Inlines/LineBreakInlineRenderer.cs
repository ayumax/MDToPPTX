using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for a <see cref="LineBreakInline"/>.
    /// </summary>
    public class LineBreakInlineRenderer : PPTXObjectRenderer<LineBreakInline>
    {
        /// <summary>
        /// Gets or sets a value indicating whether to render this softline break as a PPTX hardline break tag (&lt;br /&gt;)
        /// </summary>
        public bool RenderAsHardlineBreak { get; set; }

        protected override void Write(PPTXRenderer renderer, LineBreakInline obj)
        {
            if (obj.IsHard)
            {
                renderer.Write(obj.IsBackslash ? "\\" : "  ");
            }
            renderer.WriteLine();
        }
    }
}