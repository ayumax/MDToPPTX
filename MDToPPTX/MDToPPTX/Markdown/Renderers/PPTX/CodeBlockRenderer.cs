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
            var fencedCodeBlock = obj as FencedCodeBlock;
            if (fencedCodeBlock != null)
            {
                var opening = new string(fencedCodeBlock.FencedChar, fencedCodeBlock.FencedCharCount);
                renderer.Write(opening);
                if (fencedCodeBlock.Info != null)
                {
                    renderer.Write(fencedCodeBlock.Info);
                }
                if (!string.IsNullOrEmpty(fencedCodeBlock.Arguments))
                {
                    renderer.Write(" ").Write(fencedCodeBlock.Arguments);
                }

                renderer.WriteLine();

                renderer.WriteLeafRawLines(obj, true);
                renderer.Write(opening);
            }
            else
            {
                renderer.WriteLeafRawLines(obj, false, true);
            }

            renderer.FinishBlock();
        }
    }
}