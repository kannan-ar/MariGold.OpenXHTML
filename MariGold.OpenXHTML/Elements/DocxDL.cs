namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxDL : DocxElement
    {
        private const string defaultDDLeftMargin = "40px";
        private const string defaultDLMargin = "1em";

        private void ProcessChild(DocxProperties properties)
        {
            if (properties.CurrentNode == null)
            {
                return;
            }

            Paragraph paragraph = properties.Parent.AppendChild(new Paragraph());
            ParagraphCreated(properties.CurrentNode, paragraph);

            foreach (IHtmlNode child in properties.CurrentNode.Children)
            {
                if (child.IsText && !IsEmptyText(child.InnerHtml))
                {
                    Run run = paragraph.AppendChild(new Run(new Text()
                    {
                        Text = ClearHtml(child.InnerHtml),
                        Space = SpaceProcessingModeValues.Preserve
                    }));

                    RunCreated(properties.CurrentNode, run);
                }
                else
                {
                    ProcessChild(new DocxProperties(child, properties.CurrentNode, properties.Parent), ref paragraph);
                }
            }
        }

        private void SetDDProperties(IHtmlNode node)
        {
            DocxMargin margin = new DocxMargin(node);

            string leftMargin = margin.GetLeftMargin();

            if (string.IsNullOrEmpty(leftMargin))
            {
                //Default left margin of dd element
                margin.SetLeftMargin(defaultDDLeftMargin);
            }
        }

        private void SetMarginTop(OpenXmlElement parent)
        {
            Paragraph para = parent.AppendChild(new Paragraph());
            para.ParagraphProperties = new ParagraphProperties();

            DocxMargin.SetTopMargin(defaultDLMargin, para.ParagraphProperties);
        }

        private void SetMarginBottom(OpenXmlElement parent)
        {
            Paragraph para = parent.AppendChild(new Paragraph());
            para.ParagraphProperties = new ParagraphProperties();

            DocxMargin.SetBottomMargin(defaultDLMargin, para.ParagraphProperties);
        }

        internal DocxDL(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(IHtmlNode node)
        {
            return string.Compare(node.Tag, "dl", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxProperties properties, ref Paragraph paragraph)
        {
            if (properties.CurrentNode == null || properties.Parent == null || 
                !CanConvert(properties.CurrentNode) || IsHidden(properties.CurrentNode))
            {
                return;
            }

            if (!properties.CurrentNode.HasChildren)
            {
                return;
            }

            paragraph = null;

            //Add an empty paragraph to set default margin top
            SetMarginTop(properties.Parent);

            foreach (IHtmlNode child in properties.CurrentNode.Children)
            {
                if (string.Compare(child.Tag, "dt", StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    ProcessChild(new DocxProperties(child, properties.Parent));
                }
                else
                    if (string.Compare(child.Tag, "dd", StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        SetDDProperties(child);
                        ProcessChild(new DocxProperties(child, properties.Parent));
                    }
            }

            //Add an empty paragraph at the end to set default margin bottom
            SetMarginBottom(properties.Parent);
        }
    }
}
