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