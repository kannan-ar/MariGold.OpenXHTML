namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxHeader : DocxElement, ITextElement
    {
        private Paragraph CreateParagraph(DocxNode node)
        {
            Paragraph para = node.Parent.AppendChild(new Paragraph());
            OnParagraphCreated(node, para);
            return para;
        }

        internal DocxHeader(IOpenXmlContext context) : base(context) { }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "header", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            paragraph = null;
            Paragraph headerParagraph = null;

            foreach (DocxNode child in node.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (headerParagraph == null)
                        {
                            headerParagraph = CreateParagraph(node);
                        }

                        Run run = headerParagraph.AppendChild(new Run(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        }));

                        RunCreated(node, run);
                    }
                }
                else
                {
                    child.ParagraphNode = node;
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessChild(child, ref headerParagraph);
                }
            }
        }

        bool ITextElement.CanConvert(DocxNode node)
        {
            return CanConvert(node);
        }

        void ITextElement.Process(DocxNode node)
        {
            if (IsHidden(node))
            {
                return;
            }

            ProcessTextChild(node);
        }
    }
}
