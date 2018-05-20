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
            //var quoteIndent = renderer.Options.SpaceAfterQuoteBlock ? obj.QuoteChar + " " : obj.QuoteChar.ToString();
            //renderer.PushIndent(quoteIndent);
            renderer.WriteChildren(obj);
            //renderer.PopIndent();

            renderer.FinishBlock();
        }
    }
}