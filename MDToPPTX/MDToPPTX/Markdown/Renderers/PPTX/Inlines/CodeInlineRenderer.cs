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
            var delimiter = obj.Content.Contains(obj.Delimiter + "") ? new string(obj.Delimiter, 2) : obj.Delimiter + "";

            renderer.Write(delimiter);
            renderer.Write(obj.Content);
            renderer.Write(delimiter);
        }
    }
}