namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxBold : DocxElement, ITextElement
    {
        private void SetStyle(IHtmlNode node)
        {
            DocxNode docxNode = new DocxNode(node);

            string value = docxNode.ExtractStyleValue(DocxFont.fontWeight);

            if (string.IsNullOrEmpty(value))
            {
                docxNode.SetStyleValue(DocxFont.fontWeight, DocxFont.bold);
            }
        }

        private void ProcessRun(Run run, IHtmlNode child)
        {
            RunCreated(child, run);

            //Need to analyze the child style properties. If there is a font-weight:normal property, 
            //apply bold should not happen
            if (run.RunProperties == null)
            {
                run.RunProperties = new RunProperties();
            }

            DocxFont.ApplyBold(run.RunProperties);

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

        internal override bool CanConvert(IHtmlNode node)
        {
            return string.Compare(node.Tag, "b", StringComparison.InvariantCultureIgnoreCase) == 0 ||
            string.Compare(node.Tag, "strong", StringComparison.InvariantCultureIgnoreCase) == 0;
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

                        Run run = paragraph.AppendChild(new Run());
                        ProcessRun(run, child);
                    }
                }
                else
                {
                    SetStyle(child);
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
                    Run run = properties.Parent.AppendChild(new Run());
                    ProcessRun(run, child);
                }
                else
                {
                    SetStyle(child);
                    ProcessTextElement(new DocxProperties(child, properties.Parent));
                }
            }
        }
    }
}
