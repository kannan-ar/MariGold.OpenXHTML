namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxFooter : DocxElement, ITextElement
    {
        private Paragraph CreateParagraph(DocxProperties properties)
        {
            Paragraph para = properties.Parent.AppendChild(new Paragraph());
            ParagraphCreated(properties.CurrentNode, para);
            return para;
        }

        internal DocxFooter(IOpenXmlContext context) : base(context) { }

        internal override bool CanConvert(IHtmlNode node)
        {
            return string.Compare(node.Tag, "footer", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxProperties properties, ref Paragraph paragraph)
        {
            if (properties.CurrentNode == null || properties.Parent == null)
            {
                return;
            }

            paragraph = null;
            Paragraph footerParagraph = null;

            foreach (IHtmlNode child in properties.CurrentNode.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (footerParagraph == null)
                        {
                            footerParagraph = CreateParagraph(properties);
                        }

                        Run run = footerParagraph.AppendChild(new Run(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        }));

                        RunCreated(child, run);
                    }
                }
                else
                {
                    ProcessChild(new DocxProperties(child, properties.CurrentNode, properties.Parent), ref footerParagraph);
                }
            }
        }

        bool ITextElement.CanConvert(IHtmlNode node)
        {
            return CanConvert(node);
        }

        void ITextElement.Process(DocxProperties properties)
        {
            ProcessTextChild(properties);
        }
    }
}
