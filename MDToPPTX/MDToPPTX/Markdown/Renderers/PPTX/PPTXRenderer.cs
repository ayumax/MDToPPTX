using Markdig.Helpers;
using Markdig.Renderers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using MDToPPTX.Markdown.Renderers.PPTX.Inlines;
using MDToPPTX.PPTX;
using System;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    /// <summary>
    /// Default PPTX renderer for a Markdown <see cref="MarkdownDocument"/> object.
    /// </summary>
    public class PPTXRenderer : RendererBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PPTXRenderer"/> class.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="options">The PPTX options</param>
        public PPTXRenderer(SlideManager Writer, PPTXSetting options = null) 
        {
            this.Writer = Writer;

            Options = options ?? new PPTXSetting();
            // Default block renderers
            ObjectRenderers.Add(new CodeBlockRenderer());
            ObjectRenderers.Add(new ListRenderer());
            ObjectRenderers.Add(new HeadingRenderer());
            ObjectRenderers.Add(new HtmlBlockRenderer());
            ObjectRenderers.Add(new ParagraphRenderer());
            ObjectRenderers.Add(new QuoteBlockRenderer());
            ObjectRenderers.Add(new ThematicBreakRenderer());
            ObjectRenderers.Add(new LinkReferenceDefinitionGroupRenderer());
            ObjectRenderers.Add(new LinkReferenceDefinitionRenderer());
            ObjectRenderers.Add(new TableRenderer());

            // Default inline renderers
            ObjectRenderers.Add(new AutolinkInlineRenderer());
            ObjectRenderers.Add(new CodeInlineRenderer());
            ObjectRenderers.Add(new DelimiterInlineRenderer());
            ObjectRenderers.Add(new EmphasisInlineRenderer());
            ObjectRenderers.Add(new LineBreakInlineRenderer());
            ObjectRenderers.Add(new PPTXHtmlInlineRenderer());
            ObjectRenderers.Add(new PPTXHtmlEntityInlineRenderer());            
            ObjectRenderers.Add(new LinkInlineRenderer());
            ObjectRenderers.Add(new LiteralInlineRenderer());
        }

        public PPTXSetting Options { get; }

        private SlideManager writer;

        /// <summary>
        /// Initializes a new instance of the <see cref="PPTXRenderer"/> class.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <exception cref="System.ArgumentNullException"></exception>
        protected PPTXRenderer(SlideManager writer)
        {
            if (writer == null) throw new ArgumentNullException(nameof(writer));
            this.Writer = writer;
        }

        /// <summary>
        /// Gets or sets the writer.
        /// </summary>
        /// <exception cref="System.ArgumentNullException">if the value is null</exception>
        public SlideManager Writer
        {
            get { return writer; }
            set
            {
                writer = value ?? throw new ArgumentNullException(nameof(value));
            }
        }
        /// <summary>
        /// Renders the specified markdown object (returns the <see cref="Writer"/> as a render object).
        /// </summary>
        /// <param name="markdownObject">The markdown object.</param>
        /// <returns></returns>
        public override object Render(MarkdownObject markdownObject)
        {
            Write(markdownObject);
            return Writer;
        }

        /// <summary>
        /// Writes the inlines of a leaf inline.
        /// </summary>
        /// <param name="leafBlock">The leaf block.</param>
        /// <returns>This instance</returns>
        public PPTXRenderer WriteLeafInline(LeafBlock leafBlock)
        {
            if (leafBlock == null) throw new ArgumentNullException(nameof(leafBlock));
            var inline = (Inline)leafBlock.Inline;
            if (inline != null)
            {
                while (inline != null)
                {
                    Write(inline);
                    inline = inline.NextSibling;
                }
            }
            return this;
        }

        /// <summary>
        /// Writes the lines of a <see cref="LeafBlock"/>
        /// </summary>
        /// <param name="leafBlock">The leaf block.</param>
        /// <param name="writeEndOfLines">if set to <c>true</c> write end of lines.</param>
        /// <returns>This instance</returns>
        public PPTXRenderer WriteLeafRawLines(LeafBlock leafBlock)
        {
            if (leafBlock == null) throw new ArgumentNullException(nameof(leafBlock));
            if (leafBlock.Lines.Lines != null)
            {
                var lines = leafBlock.Lines;
                var slices = lines.Lines;
                for (int i = 0; i < lines.Count; i++)
                {
                    Write(ref slices[i].Slice);

                    WriteReturn();
                }
            }
            return this;
        }

        public void InsertNewPage()
        {
            Writer.CreateNewSlide();
        }

        /// <summary>
        /// Writes the specified content.
        /// </summary>
        /// <param name="content">The content.</param>
        /// <returns>This instance</returns>
        public PPTXRenderer Write(string content)
        {
            Writer.Write(new PPTXTextRun()
            {
                Text = content
            });
            return this;
        }

        public PPTXRenderer Write(char content)
        {
            Writer.Write(new PPTXTextRun()
            {
                Text = new string(content, 1)
            });
            return this;
        }

        public PPTXRenderer Write(PPTXTextRun content)
        {
            Writer.Write(content);
            return this;
        }

        /// <summary>
        /// Writes the specified slice.
        /// </summary>
        /// <param name="slice">The slice.</param>
        /// <returns>This instance</returns>
        public PPTXRenderer Write(ref StringSlice slice)
        {
            if (slice.Start > slice.End)
            {
                return this;
            }
            return Write(slice.Text.Substring(slice.Start, slice.Length));
        }

        /// <summary>
        /// Writes the specified slice.
        /// </summary>
        /// <param name="slice">The slice.</param>
        /// <returns>This instance</returns>
        public PPTXRenderer Write(StringSlice slice)
        {
            return Write(ref slice);
        }

        public void WriteImage(PPTXImage Image)
        {
            Writer.WriteImage(Image);
        }

        public void WriteReturn()
        {
            Writer.WriteReturn();
        }

        public PPTXTextArea StartTextArea()
        {
            return Writer.AddTextArea();
        }

        public void EndTextArea()
        {
            Writer.EndTextArea();
        }

        public void AddTextRow(PPTXText TextRow)
        {
            Writer.AddTextRow(TextRow);
        }

        public void PushFont(PPTXFont Font)
        {
            Writer.PushFont(Font);
        }

        public void PopFont()
        {
            Writer.PopFont();
        }

        public void PushHyperLink(PPTXLink Link)
        {
            Writer.PushHyperLink(Link);
        }

        public void PopHyperLink()
        {
            Writer.PopHyperLink();
        }

        public void AddTable(PPTXTable Table)
        {
            Writer.AddTable(Table);
        }

        public void AddTableEnd()
        {
            Writer.AddTableEnd();
        }


        public void AddTableRow()
        {
            Writer.AddTableRow();
        }

        public void NextTableCell()
        {
            Writer.NextTableCell();
        }

        public void EndTableRow()
        {
            Writer.EndTableRow();
        }

    }
}