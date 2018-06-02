using Markdig.Syntax;
using MDToPPTX.PPTX;

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
            renderer.PushBlockSetting(renderer.Options.Code);

            renderer.Write(" ");
            renderer.WriteReturn();
            renderer.WriteLeafRawLines(obj);
            renderer.WriteReturn();
            renderer.Write(" ");
           
            renderer.PopBlockSetting();

            renderer.EndTextArea();
        }
    }
}