using Markdig.Syntax;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    /// <summary>
    /// An PPTX renderer for a <see cref="CodeBlock"/> and <see cref="FencedCodeBlock"/>.
    /// </summary>
    public class CodeBlockRenderer : PPTXObjectRenderer<CodeBlock>
    {
        public bool OutputAttributesOnPre { get; set; }

        protected override void Write(PPTXRenderer renderer, CodeBlock obj)
        {
            renderer.StartTextArea();

            renderer.PushFont(renderer.Options.CodeFont);
            renderer.WriteLeafRawLines(obj);
            renderer.PopFont();

            renderer.EndTextArea();
        }
    }
}