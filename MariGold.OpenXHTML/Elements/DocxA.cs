namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxA : DocxElement
    {
        private const string href = "href";

        private void CreateParagraph(DocxProperties properties, ref Paragraph paragraph)
        {
            if (paragraph == null)
            {
                paragraph = properties.Parent.AppendChild(new Paragraph());
                ParagraphCreated(properties.ParagraphNode, paragraph);
            }
        }

        private void ProcessNonLinkText(DocxProperties properties, ref Paragraph paragraph)
        {
            foreach (IHtmlNode child in properties.CurrentNode.Children)
            {
                if (child.IsText)
                {
                    if (paragraph == null)
                    {
                        paragraph = properties.Parent.AppendChild(new Paragraph());
                        ParagraphCreated(properties.ParagraphNode, paragraph);
                    }

                    if (!IsEmptyText(child.InnerHtml))
                    {
                        paragraph.AppendChild<Run>(new Run(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        }));
                    }
                }
                else
                {
                    ProcessChild(new DocxProperties(child, properties.ParagraphNode, properties.Parent), ref paragraph);
                }
            }
        }

        private void ProcessChildren(DocxProperties currentProperties, DocxProperties newProperties, Run run)
        {
            foreach (IHtmlNode child in currentProperties.CurrentNode.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        run.AppendChild(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        });
                    }
                }
                else
                {
                    ProcessTextElement(newProperties);
                }
            }
        }

        internal DocxA(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(IHtmlNode node)
        {
            return string.Compare(node.Tag, "a", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxProperties properties, ref Paragraph paragraph)
        {
            if (properties.CurrentNode == null)
            {
                return;
            }

            DocxNode docxNode = new DocxNode(properties.CurrentNode);

            string link = docxNode.ExtractAttributeValue(href);

            link = CleanUrl(link);

            if (Uri.IsWellFormedUriString(link, UriKind.Absolute))
            {
                Uri uri = new Uri(link);

                var relationship = context.MainDocumentPart.AddHyperlinkRelationship(uri, uri.IsAbsoluteUri);

                var hyperLink = new Hyperlink() { History = true, Id = relationship.Id };

                foreach (IHtmlNode child in properties.CurrentNode.Children)
                {
                    if (child.IsText)
                    {
                        if (!IsEmptyText(child.InnerHtml))
                        {
                            Run run = hyperLink.AppendChild<Run>(new Run(new Text()
                             {
                                 Text = ClearHtml(child.InnerHtml),
                                 Space = SpaceProcessingModeValues.Preserve
                             }));

                            RunCreated(properties.CurrentNode, run);

                            if (run.RunProperties == null)
                            {
                                run.RunProperties = new RunProperties((new RunStyle() { Val = "Hyperlink" }));
                            }
                            else
                            {
                                run.RunProperties.Append(new RunStyle() { Val = "Hyperlink" });
                            }
                        }
                    }
                    else
                    {
                        ProcessTextElement(new DocxProperties(child, hyperLink));
                    }
                }

                CreateParagraph(properties, ref paragraph);

                paragraph.Append(hyperLink);
            }
            else
            {
                ProcessNonLinkText(properties, ref paragraph);
            }
        }
    }
}
