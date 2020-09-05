namespace MariGold.OpenXHTML
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxDL : DocxElement
    {
        private const string defaultDDLeftMargin = "40px";
        private const string defaultDLMargin = "1em";

        private void ProcessChild(DocxNode node, Dictionary<string, object> properties)
        {
            if (node.IsNull())
            {
                return;
            }

            Paragraph paragraph = node.Parent.AppendChild(new Paragraph());
            OnParagraphCreated(node, paragraph);

            foreach (DocxNode child in node.Children)
            {
                if (child.IsText && !IsEmptyText(child.InnerHtml))
                {
                    Run run = paragraph.AppendChild(new Run(new Text()
                    {
                        Text = ClearHtml(child.InnerHtml),
                        Space = SpaceProcessingModeValues.Preserve
                    }));

                    RunCreated(node, run);
                }
                else
                {
                    child.ParagraphNode = node;
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessChild(child, ref paragraph, properties);
                }
            }
        }

        private void SetDDProperties(DocxNode node)
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

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "dl", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (node.IsNull() || node.Parent == null || !CanConvert(node) || IsHidden(node))
            {
                return;
            }

            if (!node.HasChildren)
            {
                return;
            }

            paragraph = null;

            //Add an empty paragraph to set default margin top
            SetMarginTop(node.Parent);

            foreach (DocxNode child in node.Children)
            {
                if (string.Compare(child.Tag, "dt", StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessChild(child, properties);
                }
                else if (string.Compare(child.Tag, "dd", StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    node.CopyExtentedStyles(child);
                    SetDDProperties(child);
                    child.Parent = node.Parent;
                    ProcessChild(child, properties);
                }
            }

            //Add an empty paragraph at the end to set default margin bottom
            SetMarginBottom(node.Parent);
        }
    }
}
