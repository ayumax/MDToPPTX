using System.Globalization;
using Markdig.Syntax;
using MDToPPTX.PPTX;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    /// <summary>
    /// A PPTX renderer for a <see cref="ListBlock"/>.
    /// </summary>
    public class ListRenderer : PPTXObjectRenderer<ListBlock>
    {
        protected override void Write(PPTXRenderer renderer, ListBlock listBlock)
        {
            renderer.StartTextArea();

            for (var i = 0; i < listBlock.Count; i++)
            {
                var item = listBlock[i];
                var listItem = (ListItemBlock)item;

                renderer.AddTextRow(new PPTXText()
                {
                    Bullet = listBlock.IsOrdered ? PPTXBullet.Number : PPTXBullet.Circle
                });

                renderer.WriteChildren(listItem);

                renderer.WriteReturn();
            }

            renderer.EndTextArea();
        }

    }
}