using Markdig.Syntax;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    public class LinkReferenceDefinitionRenderer : PPTXObjectRenderer<LinkReferenceDefinition>
    {
        protected override void Write(PPTXRenderer renderer, LinkReferenceDefinition linkDef)
        {
            renderer.StartTextArea();
            renderer.Write('[');            
            renderer.Write(linkDef.Label);
            renderer.Write("]: ");

            renderer.Write(linkDef.Url);

            if (linkDef.Title != null)
            {
                renderer.Write(" \"");
                renderer.Write(linkDef.Title.Replace("\"", "\\\""));
                renderer.Write('"');
            }
            renderer.EndTextArea();
        }
    }
}