namespace MariGold.OpenXHTML
{
    using System;
    using System.Collections.Generic;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.Text.RegularExpressions;

    internal sealed class DocxHeading : DocxElement
    {
        private Regex isValid;

        private int GetHeaderNumber(DocxNode node)
        {
            int value = -1;
            Regex regex = new Regex("[1-6]{1}$");

            Match match = regex.Match(node.Tag);

            if (match != null)
            {
                Int32.TryParse(match.Value, out value);
            }

            return value;
        }

        private string CalculateFontSize(int headerSize)
        {
            string fontSize = string.Empty;

            switch (headerSize)
            {
                case 1:
                    fontSize = "2em";
                    break;

                case 2:
                    fontSize = "1.5em";
                    break;

                case 3:
                    fontSize = "1.17em";
                    break;

                case 4:
                    fontSize = "1em";
                    break;

                case 5:
                    fontSize = ".83em";
                    break;

                case 6:
                    fontSize = ".67em";
                    break;
            }

            return fontSize;
        }

        private void ApplyStyle(IHtmlNode node)
        {
            DocxNode docxNode = new DocxNode(node);
            string fontSizeValue = docxNode.ExtractStyleValue(DocxFont.fontSize);
            string fontWeightValue = docxNode.ExtractStyleValue(DocxFont.fontWeight);

            if (string.IsNullOrEmpty(fontSizeValue))
            {
                fontSizeValue = CalculateFontSize(GetHeaderNumber(docxNode));
            }

            if(string.IsNullOrEmpty(fontWeightValue))
            {
                fontWeightValue = DocxFont.bold;
            }

            Dictionary<string, string> newStyles = new Dictionary<string, string>();

            newStyles.Add(DocxFont.fontSize, fontSizeValue);
            newStyles.Add(DocxFont.fontWeight, fontWeightValue);

            docxNode.SetStyleValues(newStyles);
        }

        internal DocxHeading(IOpenXmlContext context)
            : base(context)
        {
            isValid = new Regex(@"^[hH][1-6]{1}$");
        }

        internal override bool CanConvert(IHtmlNode node)
        {
            return isValid.IsMatch(node.Tag);
        }

        internal override void Process(DocxProperties properties, ref Paragraph paragraph)
        {
            if (properties.CurrentNode == null || properties.Parent == null
                || IsHidden(properties.CurrentNode))
            {
                return;
            }

            paragraph = null;
            Paragraph headerParagraph = null;

            foreach (IHtmlNode child in properties.CurrentNode.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (headerParagraph == null)
                        {
                            headerParagraph = properties.Parent.AppendChild(new Paragraph());
                            ParagraphCreated(properties.CurrentNode, headerParagraph);
                        }

                        Run run = headerParagraph.AppendChild(new Run());
                        ApplyStyle(properties.CurrentNode);
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
                    ProcessChild(new DocxProperties(child, properties.CurrentNode, properties.Parent), ref headerParagraph);
                }
            }
        }
    }
}
