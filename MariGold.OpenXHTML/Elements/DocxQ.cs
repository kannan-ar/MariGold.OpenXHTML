namespace MariGold.OpenXHTML
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxQ : DocxElement, ITextElement
    {
        private bool hasOpenQuote;

        private Paragraph CreateParagraph(DocxNode node)
        {
            Paragraph paragraph = node.Parent.AppendChild(new Paragraph());
            OnParagraphCreated(node.ParagraphNode, paragraph);
           
            return paragraph;
        }

        private void ApplyOpenQuoteIfEmpty(DocxNode node, ref Paragraph paragraph)
        {
            if (paragraph == null)
            {
                paragraph = CreateParagraph(node);
            }

            if (hasOpenQuote)
            {
                return;
            }

            paragraph.AppendChild(new Run(new Text() { Text = "\"" }));
            hasOpenQuote = true;
        }

        internal DocxQ(IOpenXmlContext context) : base(context) { }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "q", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            foreach (DocxNode child in node.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        ApplyOpenQuoteIfEmpty(node, ref paragraph);

                        Run run = paragraph.AppendChild(new Run(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        }));

                        RunCreated(node, run);
                    }
                }
                else
                {
                    child.ParagraphNode = node.ParagraphNode;
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ApplyOpenQuoteIfEmpty(node, ref paragraph);
                    ProcessChild(child, ref paragraph, properties);
                }
            }

            if (paragraph != null && hasOpenQuote)
            {
                paragraph.AppendChild(new Run(new Text() { Text = "\"" }));
            }
        }

        bool ITextElement.CanConvert(DocxNode node)
        {
            return CanConvert(node);
        }

        void ITextElement.Process(DocxNode node, Dictionary<string, object> properties)
        {
            if (IsHidden(node))
            {
                return;
            }

            ProcessTextChild(node, properties);
        }
    }
}
