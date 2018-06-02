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
                renderer.PushBlockSetting(renderer.Options.Normal);
            }

            renderer.WriteLeafInline(obj);

            if (obj.Parent is MarkdownDocument)
            {
                renderer.PopBlockSetting();
                renderer.EndTextArea();
            }
        }
    }
}