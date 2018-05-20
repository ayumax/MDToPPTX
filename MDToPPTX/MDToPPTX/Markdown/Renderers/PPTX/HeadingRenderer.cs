using Markdig.Syntax;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    /// <summary>
    /// An PPTX renderer for a <see cref="HeadingBlock"/>.
    /// </summary>
    public class HeadingRenderer : PPTXObjectRenderer<HeadingBlock>
    {
        private static readonly string[] HeadingTexts = {
            "#",
            "##",
            "###",
            "####",
            "#####",
            "######",
        };

        protected override void Write(PPTXRenderer renderer, HeadingBlock obj)
        {
            var headingText = obj.Level > 0 && obj.Level <= 6
                ? HeadingTexts[obj.Level - 1]
                : new string('#', obj.Level);

            renderer.Write(headingText);
            renderer.WriteLeafInline(obj);

            renderer.FinishBlock();
        }
    }
}