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
            if (obj.Parent is MarkdownDocument)
            {
                renderer.PushFont(renderer.Options.NormalFont);
            }

            renderer.WriteLeafInline(obj);

            if (obj.Parent is MarkdownDocument)
            {
                renderer.PopFont();
                renderer.EndTextArea();
            }
        }
    }
}