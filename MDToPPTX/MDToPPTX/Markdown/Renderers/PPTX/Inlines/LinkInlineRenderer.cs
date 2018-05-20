using Markdig.Syntax.Inlines;

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
                renderer.Write('!');
            }
            renderer.Write('[');
            renderer.WriteChildren(link);
            renderer.Write(']');

            if (link.Label != null)
            {

                var literal = link.FirstChild as LiteralInline;
                if (literal != null && literal.Content.Match(link.Label) && literal.Content.Length == link.Label.Length)
                {
                    // collapsed reference and shortcut links
                    if (!link.IsShortcut)
                    {
                        renderer.Write("[]");
                    }
                }
                else
                {
                    // full link
                    renderer.Write('[').Write(link.Label).Write(']');
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(link.Url))
                {
                    renderer.Write('(').Write(link.Url);

                    if (!string.IsNullOrEmpty(link.Title))
                    {
                        renderer.Write(" \"");
                        renderer.Write(link.Title.Replace(@"""", @"\"""));
                        renderer.Write("\"");
                    }

                    renderer.Write(')');
                }
            }
        }
    }
}