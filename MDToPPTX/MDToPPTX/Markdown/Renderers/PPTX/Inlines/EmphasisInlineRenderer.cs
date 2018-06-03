using Markdig.Syntax.Inlines;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for an <see cref="EmphasisInline"/>.
    /// </summary>
    public class EmphasisInlineRenderer : PPTXObjectRenderer<EmphasisInline>
    {
        protected override void Write(PPTXRenderer renderer, EmphasisInline obj)
        {         
            var cloneFont = renderer.Writer.CurrentFont.Clone();
            
            if (obj.DelimiterChar == '*' || obj.DelimiterChar == '_')
            {
                if (obj.IsDouble)
                {
                    cloneFont.Bold = true;
                }
                else
                {
                    cloneFont.Italic = true;
                }
            }
            else if (obj.DelimiterChar == '~')
            {
                if (obj.IsDouble)
                {
                    cloneFont.Strike = true;
                }
            }

            renderer.Writer.FontStack.Push(cloneFont);

            renderer.WriteChildren(obj);

            renderer.Writer.FontStack.Pop();
        }
    }
}