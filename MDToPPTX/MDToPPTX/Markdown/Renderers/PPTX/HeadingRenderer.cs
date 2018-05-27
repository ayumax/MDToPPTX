using Markdig.Syntax;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    /// <summary>
    /// An PPTX renderer for a <see cref="HeadingBlock"/>.
    /// </summary>
    public class HeadingRenderer : PPTXObjectRenderer<HeadingBlock>
    {
        protected override void Write(PPTXRenderer renderer, HeadingBlock obj)
        {
            renderer.StartTextArea();

            var setFont = renderer.Options.NormalFont;
            switch (obj.Level)
            {
                case 1:
                    setFont = renderer.Options.Header1font;
                    break;
                case 2:
                    setFont = renderer.Options.Header2font;
                    break;
            }

            renderer.PushFont(setFont);
            renderer.WriteLeafInline(obj);
            renderer.PopFont();

            renderer.EndTextArea();
        }
    }
}