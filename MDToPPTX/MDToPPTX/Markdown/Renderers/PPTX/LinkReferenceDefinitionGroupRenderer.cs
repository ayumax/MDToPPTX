using Markdig.Syntax;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    public class LinkReferenceDefinitionGroupRenderer : PPTXObjectRenderer<LinkReferenceDefinitionGroup>
    {
        protected override void Write(PPTXRenderer renderer, LinkReferenceDefinitionGroup obj)
        {
            renderer.StartTextArea();
            renderer.WriteChildren(obj);
            renderer.EndTextArea();
        }
    }
}