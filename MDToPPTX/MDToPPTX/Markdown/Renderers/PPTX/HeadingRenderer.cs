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
            var _block = renderer.Options.Normal;
            switch (obj.Level)
            {
                case 1:
                    _block = renderer.Options.Header1;
                    break;
                case 2:
                    _block = renderer.Options.Header2;
                    break;
            }

            renderer.PushBlockSetting(_block);

            renderer.StartTextArea();

            renderer.WriteLeafInline(obj);
            renderer.PopBlockSetting();

            renderer.EndTextArea();
        }
    }
}