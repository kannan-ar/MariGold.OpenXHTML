namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxItalic : DocxElement, ITextElement
    {
        private void SetStyle(DocxNode node)
        {
            string value = node.ExtractStyleValue(DocxFontStyle.italic);

            if (string.IsNullOrEmpty(value))
            {
                node.SetExtentedStyle(DocxFontStyle.fontStyle, DocxFontStyle.italic);
            }
        }

        internal DocxItalic(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "i", StringComparison.InvariantCultureIgnoreCase) == 0 ||
            string.Compare(node.Tag, "em", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || IsHidden(node))
            {
                return;
            }

            SetStyle(node);

            ProcessElement(node, ref paragraph);

            /*
            foreach (DocxNode child in node.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (paragraph == null)
                        {
                            paragraph = node.Parent.AppendChild(new Paragraph());
                            OnParagraphCreated(node.ParagraphNode, paragraph);
                        }

                        Run run = paragraph.AppendChild(new Run());
                        RunCreated(node, run);
                        
                        run.AppendChild(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        });
                    }
                }
                else
                {
                    child.ParagraphNode = node.ParagraphNode;
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessChild(child, ref paragraph);
                }
            }
             * */
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

            SetStyle(node);
            ProcessTextChild(node);
        }
    }
}
