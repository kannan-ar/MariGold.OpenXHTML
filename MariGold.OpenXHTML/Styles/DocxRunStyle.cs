﻿namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxRunStyle
    {
        private void CheckFonts(DocxNode node, RunProperties properties)
        {
            string fontFamily = node.ExtractStyleValue(DocxFontStyle.fontFamily);
            string fontWeight = node.ExtractStyleValue(DocxFontStyle.fontWeight);
            string fontStyle = node.ExtractStyleValue(DocxFontStyle.fontStyle);

            if (!string.IsNullOrEmpty(fontFamily))
            {
                DocxFontStyle.ApplyFontFamily(fontFamily, properties);
            }

            if (!string.IsNullOrEmpty(fontWeight))
            {
                DocxFontStyle.ApplyFontWeight(fontWeight, properties);
            }

            if (!string.IsNullOrEmpty(fontStyle))
            {
                DocxFontStyle.ApplyFontStyle(fontStyle, properties);
            }
        }

        private void CheckFontStyle(DocxNode node, RunProperties properties)
        {
            string fontSize = node.ExtractStyleValue(DocxFontStyle.fontSize);
            string textDecoration = node.ExtractStyleValue(DocxFontStyle.textDecoration);

            if (string.IsNullOrEmpty(textDecoration))
            {
                textDecoration = node.ExtractStyleValue(DocxFontStyle.textDecorationLine);
            }

            if (!string.IsNullOrEmpty(fontSize))
            {
                DocxFontStyle.ApplyFontSize(fontSize, properties);
            }

            if (!string.IsNullOrEmpty(textDecoration))
            {
                DocxFontStyle.ApplyTextDecoration(textDecoration, properties);
            }
        }

        private void ProcessBackGround(DocxNode node, RunProperties properties)
        {
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
        }

        private void ProcessVerticalAlign(DocxNode node, RunProperties properties)
        {
            string verticalAlign = node.ExtractStyleValue(DocxAlignment.verticalAlign);

            if (!string.IsNullOrEmpty(verticalAlign))
            {
                DocxAlignment.ApplyVerticalTextAlign(verticalAlign, properties);
            }
        }

        private void ProcessDirection(DocxNode node, RunProperties properties)
        {
            string styleValue = node.ExtractStyleValue(DocxDirection.direction);

            if (!string.IsNullOrEmpty(styleValue))
            {
                DocxDirection.ApplyDirection(styleValue, properties);
            }
        }

        internal void Process(Run element, DocxNode node)
        {
            RunProperties properties = element.RunProperties;

            if (properties == null)
            {
                properties = new RunProperties();
            }

            //Order of assigning styles to run property is important. The order should not change.
            CheckFonts(node, properties);

            string color = node.ExtractStyleValue(DocxColor.color);

            if (!string.IsNullOrEmpty(color))
            {
                DocxColor.ApplyColor(color, properties);
            }

            CheckFontStyle(node, properties);

            ProcessBackGround(node, properties);

            ProcessVerticalAlign(node, properties);

            ProcessDirection(node, properties);

            if (element.RunProperties == null && properties.HasChildren)
            {
                element.RunProperties = properties;
            }
        }
    }
}
