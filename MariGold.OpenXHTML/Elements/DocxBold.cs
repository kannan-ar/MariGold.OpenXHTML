namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxBold : DocxElement, ITextElement
    {
        private void SetStyle(DocxNode node)
        {
            string value = node.ExtractStyleValue(DocxFont.fontWeight);

            if (string.IsNullOrEmpty(value))
            {
                node.SetExtentedStyle(DocxFont.fontWeight, DocxFont.bold);
            }
        }

        private void ProcessRun(Run run, DocxNode parent, DocxNode child)
        {
            RunCreated(parent, run);

            //Need to analyze the child style properties. If there is a font-weight:normal property, 
            //apply bold should not happen
            /*
            if (run.RunProperties == null)
            {
                run.RunProperties = new RunProperties();
            }

            DocxFont.ApplyBold(run.RunProperties);
            */
            run.AppendChild(new Text()
            {
                Text = ClearHtml(child.InnerHtml),
                Space = SpaceProcessingModeValues.Preserve
            });
        }

        public DocxBold(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "b", StringComparison.InvariantCultureIgnoreCase) == 0 ||
            string.Compare(node.Tag, "strong", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || IsHidden(node))
            {
                return;
            }

            SetStyle(node);

            foreach (DocxNode child in node.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (paragraph == null)
                        {
                            paragraph = node.Parent.AppendChild(new Paragraph());
                            ParagraphCreated(node.ParagraphNode, paragraph);
                        }

                        Run run = paragraph.AppendChild(new Run());
                        ProcessRun(run, node, child);
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

            foreach (DocxNode child in node.Children)
            {
                if (child.IsText && !IsEmptyText(child.InnerHtml))
                {
                    Run run = node.Parent.AppendChild(new Run());
                    ProcessRun(run, node, child);
                }
                else
                {
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessTextElement(child);
                }
            }
        }
    }
}
