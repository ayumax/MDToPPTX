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
            var setFont = renderer.Options.Normal;
            switch (obj.Level)
            {
                case 1:
                    setFont = renderer.Options.Header1;
                    break;
                case 2:
                    setFont = renderer.Options.Header2;
                    break;
            }

            renderer.PushBlockSetting(setFont);
            renderer.WriteLeafInline(obj);
            renderer.PopBlockSetting();

            renderer.EndTextArea();
        }
    }
}