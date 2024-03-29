﻿namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections.Generic;

    internal sealed class DocxA : DocxElement
    {
        private const string href = "href";

        private void CreateParagraph(DocxNode node, ref Paragraph paragraph)
        {
            if (paragraph == null)
            {
                paragraph = node.Parent.AppendChild(new Paragraph());
                OnParagraphCreated(node.ParagraphNode, paragraph);
            }
        }

        private void ProcessNonLinkText(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            foreach (DocxNode child in node.Children)
            {
                if (child.IsText)
                {
                    if (paragraph == null)
                    {
                        paragraph = node.Parent.AppendChild(new Paragraph());
                        OnParagraphCreated(node.ParagraphNode, paragraph);
                    }

                    if (!IsEmptyText(child.InnerHtml))
                    {
                        Run run = paragraph.AppendChild<Run>(new Run(new Text()
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
                    ProcessChild(child, ref paragraph, properties);
                }
            }
        }

        internal DocxA(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "a", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (node.IsNull() || IsHidden(node))
            {
                return;
            }

            string link = node.ExtractAttributeValue(href);

            link = CleanUrl(link);

            if (Uri.IsWellFormedUriString(link, UriKind.Absolute))
            {
                Uri uri = new Uri(link);

                var relationship = context.MainDocumentPart.AddHyperlinkRelationship(uri, uri.IsAbsoluteUri);

                var hyperLink = new Hyperlink() { History = true, Id = relationship.Id };

                foreach (DocxNode child in node.Children)
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

                            run.RunProperties = new RunProperties((new RunStyle() { Val = "Hyperlink" }));
                            RunCreated(node, run);
                        }
                    }
                    else
                    {
                        child.Parent = hyperLink;
                        node.CopyExtentedStyles(child);
                        ProcessTextElement(child, properties);
                    }
                }

                CreateParagraph(node, ref paragraph);
                paragraph.Append(hyperLink);
            }
            else
            {
                ProcessNonLinkText(node, ref paragraph, properties);
            }
        }
    }
}
