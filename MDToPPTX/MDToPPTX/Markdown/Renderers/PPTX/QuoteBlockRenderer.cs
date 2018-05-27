using Markdig.Syntax;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    /// <summary>
    /// A PPTX renderer for a <see cref="QuoteBlock"/>.
    /// </summary>
    public class QuoteBlockRenderer : PPTXObjectRenderer<QuoteBlock>
    {
        protected override void Write(PPTXRenderer renderer, QuoteBlock obj)
        {
            renderer.StartTextArea();

            renderer.WriteChildren(obj);

            renderer.EndTextArea();
        }
    }
}