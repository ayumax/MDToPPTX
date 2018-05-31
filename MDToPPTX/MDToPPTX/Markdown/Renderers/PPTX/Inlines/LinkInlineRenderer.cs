using Markdig.Syntax.Inlines;
using MDToPPTX.PPTX;

namespace MDToPPTX.Markdown.Renderers.PPTX.Inlines
{
    /// <summary>
    /// A PPTX renderer for a <see cref="LinkInline"/>.
    /// </summary>
    public class LinkInlineRenderer : PPTXObjectRenderer<LinkInline>
    {
        protected override void Write(PPTXRenderer renderer, LinkInline link)
        {
            if (link.IsImage)
            {
                renderer.EndTextArea();

                WriteImageLink(renderer, link);
            }
            else
            {
                WriteHyperLink(renderer, link);
            }
        }

        private void WriteImageLink(PPTXRenderer renderer, LinkInline link)
        {
            renderer.WriteImage(new PPTXImage(link.Url));
        }

        private void WriteHyperLink(PPTXRenderer renderer, LinkInline link)
        {
            renderer.PushHyperLink(new PPTXLink()
            {
                LinkKey = link.Url,
                LinkURL = link.Url
            });

            if (link.Label != null)
            {
                var literal = link.FirstChild as LiteralInline;
                if (literal != null && literal.Content.Match(link.Label) && literal.Content.Length == link.Label.Length)
                {
                }
                else
                {
                    // full link
                    renderer.Write(link.Label);
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(link.Url))
                {
                    renderer.WriteChildren(link);    
                }
            }

            renderer.PopHyperLink();
        }
    }
}