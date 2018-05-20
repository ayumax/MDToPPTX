// Copyright (c) Alexandre Mutel. All rights reserved.
// This file is licensed under the BSD-Clause 2 license. 
// See the license.txt file in the project root for more information.
using Markdig.Syntax;
using Markdig.Renderers;

namespace MDToPPTX.Markdown.Renderers.PPTX
{
    /// <summary>
    /// A base class for PPTX rendering <see cref="Block"/> and <see cref="Markdig.Syntax.Inlines.Inline"/> Markdown objects.
    /// </summary>
    /// <typeparam name="TObject">The type of the object.</typeparam>
    /// <seealso cref="Markdig.Renderers.IMarkdownObjectRenderer" />
    public abstract class PPTXObjectRenderer<TObject> : MarkdownObjectRenderer<PPTXRenderer, TObject> where TObject : MarkdownObject
    {
    }
}