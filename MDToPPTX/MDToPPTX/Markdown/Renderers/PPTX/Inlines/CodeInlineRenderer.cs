using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for a <see cref="CodeInline"/>.
    /// </summary>
    public class CodeInlineRenderer : PPTXObjectRenderer<CodeInline>
    {
        protected override void Write(PPTXRenderer renderer, CodeInline obj)
        {
            renderer.PushInlineSetting(renderer.Options.InlineCode);
            renderer.Write(obj.Content);
            renderer.PopBlockSetting();
        }
    }
}