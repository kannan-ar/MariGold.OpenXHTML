﻿namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;

    internal sealed class DocxParagraphStyle
    {
        private const string pageBreakBefore = "page-break-before";
        private const string pageBreakAlways = "always";

        private void ProcessBorder(DocxNode node, ParagraphProperties properties)
        {
            ParagraphBorders paragraphBorders = new ParagraphBorders();

            DocxBorder.ApplyBorders(paragraphBorders,
                node.ExtractStyleValue(DocxBorder.borderName),
                node.ExtractStyleValue(DocxBorder.leftBorderName),
                node.ExtractStyleValue(DocxBorder.topBorderName),
                node.ExtractStyleValue(DocxBorder.rightBorderName),
                node.ExtractStyleValue(DocxBorder.bottomBorderName),
                false);

            if (paragraphBorders.HasChildren)
            {
                properties.Append(paragraphBorders);
            }
        }

        static internal void SetIndentation(Paragraph element, int indent)
        {
            if (element.ParagraphProperties == null)
            {
                element.ParagraphProperties = new ParagraphProperties();
            }

            element.ParagraphProperties.Append(new Indentation() { Left = indent.ToString() });
        }

        internal void Process(Paragraph element, DocxNode node)
        {
            ParagraphProperties properties = element.ParagraphProperties;

            if (properties == null)
            {
                properties = new ParagraphProperties();
            }

            var pageBreak = node.ExtractStyleValue(pageBreakBefore);

            if (!string.IsNullOrEmpty(pageBreak) && string.Compare(pageBreak, pageBreakAlways, StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                properties.PageBreakBefore = new PageBreakBefore();
            }

            //Order of assigning styles to paragraph property is important. The order should not change.
            ProcessBorder(node, properties);

            string backgroundColor = node.ExtractStyleValue(DocxColor.backGroundColor);
            string backGround = DocxColor.ExtractBackGround(node.ExtractStyleValue(DocxColor.backGround));

            if (!string.IsNullOrEmpty(backgroundColor))
            {
                DocxColor.ApplyBackGroundColor(backgroundColor, properties);
            }
            else if (!string.IsNullOrEmpty(backGround))
            {
                DocxColor.ApplyBackGroundColor(backGround, properties);
            }

            DocxMargin margin = new DocxMargin(node);
            margin.ProcessParagraphMargin(properties);

            string textAlign = node.ExtractStyleValue(DocxAlignment.textAlign);
            if (!string.IsNullOrEmpty(textAlign))
            {
                DocxAlignment.ApplyTextAlign(textAlign, properties);
            }

            if (element.ParagraphProperties == null && properties.HasChildren)
            {
                element.ParagraphProperties = properties;
            }
        }
    }
}
