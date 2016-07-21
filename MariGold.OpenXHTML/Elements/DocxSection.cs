namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxSection : DocxElement
    {
        private const string tag = "section";

        internal DocxSection(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(IHtmlNode node)
        {
            return string.Compare(node.Tag, tag, StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxProperties properties, ref Paragraph paragraph)
        {
            if (properties.CurrentNode == null || properties.Parent == null || IsHidden(properties.CurrentNode))
            {
                return;
            }

            paragraph = null;
            Paragraph sectionParagraph = null;

            foreach (IHtmlNode child in properties.CurrentNode.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (sectionParagraph == null)
                        {
                            sectionParagraph = properties.Parent.AppendChild(new Paragraph());
                            ParagraphCreated(properties.ParagraphNode, sectionParagraph);
                        }

                        Run run = sectionParagraph.AppendChild(new Run(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        }));
                        RunCreated(child, run);
                    }
                }
                else
                {
                    ProcessChild(new DocxProperties(child, properties.ParagraphNode, properties.Parent), ref sectionParagraph);
                }
            }
        }
    }
}
