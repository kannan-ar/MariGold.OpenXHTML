namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxCenter : DocxElement, ITextElement
    {
        internal DocxCenter(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(IHtmlNode node)
        {
            return string.Compare(node.Tag, "center", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxProperties properties, ref Paragraph paragraph)
        {
            if (properties.CurrentNode == null)
            {
                return;
            }

            foreach (IHtmlNode child in properties.CurrentNode.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (paragraph == null)
                        {
                            paragraph = properties.Parent.AppendChild(new Paragraph());
                            ParagraphCreated(properties.ParagraphNode, paragraph);
                        }

                        if (paragraph.ParagraphProperties == null)
                        {
                            paragraph.ParagraphProperties = new ParagraphProperties();
                        }

                        DocxAlignment.AlignCenter(paragraph.ParagraphProperties);

                        Run run = paragraph.AppendChild(new Run());
                        RunCreated(properties.CurrentNode, run);

                        run.AppendChild(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        });
                    }
                }
                else
                {
                    ProcessChild(new DocxProperties(child, properties.ParagraphNode, properties.Parent), ref paragraph);
                }
            }
        }

        bool ITextElement.CanConvert(IHtmlNode node)
        {
            return CanConvert(node);
        }

        void ITextElement.Process(DocxProperties properties)
        {
            foreach (IHtmlNode child in properties.CurrentNode.Children)
            {
                if (child.IsText && !IsEmptyText(child.InnerHtml))
                {
                    Run run = properties.Parent.AppendChild(new Run(new Text()
                    {
                        Text = ClearHtml(child.InnerHtml),
                        Space = SpaceProcessingModeValues.Preserve
                    }));

                    RunCreated(child, run);
                }
                else
                {
                    ProcessTextElement(new DocxProperties(child, properties.Parent));
                }
            }
        }
    }
}
