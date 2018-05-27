using Markdig.Syntax;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    /// <summary>
    /// A PPTX renderer for a <see cref="ParagraphBlock"/>.
    /// </summary>
    public class ParagraphRenderer : PPTXObjectRenderer<ParagraphBlock>
    {
        protected override void Write(PPTXRenderer renderer, ParagraphBlock obj)
        {
            // TopŠK‘w‚Ì‚Ì‚İƒGƒŠƒA‚ğV‹K‚Åì¬
            if (obj.Parent is MarkdownDocument)
            {
                renderer.StartTextArea();
            }

            renderer.PushFont(renderer.Options.NormalFont);

            renderer.WriteLeafInline(obj);

            renderer.PopFont();

            if (obj.Parent is MarkdownDocument)
            {
                renderer.EndTextArea();
            }
        }
    }
}