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
            var myArea = renderer.StartTextArea();
            myArea.BackgroundColor = new PPTXColor(240, 240, 240);

            renderer.PushFont(renderer.Options.CodeFont);
            renderer.Write(" ");
            renderer.WriteReturn();
            renderer.WriteLeafRawLines(obj);
            renderer.WriteReturn();
            renderer.Write(" ");
           
            renderer.PopFont();

            renderer.EndTextArea();
        }
    }
}