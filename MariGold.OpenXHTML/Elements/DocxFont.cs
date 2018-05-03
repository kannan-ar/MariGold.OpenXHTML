namespace MariGold.OpenXHTML
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;

    internal sealed class DocxFont : DocxElement, ITextElement
    {
        private const string defaultFontSize = "16px";
        private readonly Dictionary<Int32, Int32> fontSizes;

        private void SetFontSize(DocxNode node)
        {
            string size = node.ExtractAttributeValue("size");

            if (string.IsNullOrEmpty(size))
            {
                return;
            }

            Match match = Regex.Match(size, "^\\d+");
            Int32 sizeValue;
            Int32 fontSizeValue = 0;

            if (!match.Success || !Int32.TryParse(match.Value, out sizeValue))
            {
                return;
            }

            if (!fontSizes.TryGetValue(sizeValue, out fontSizeValue))
            {
                if (sizeValue > 7)
                {
                    fontSizeValue = 48;
                }
            }

            if (fontSizeValue != 0)
            {
                node.SetExtentedStyle(DocxFontStyle.fontSize, string.Concat(fontSizeValue.ToString(), "px"));
            }
        }

        private void ApplyStyle(DocxNode node)
        {
            string fontFamily = node.ExtractOwnStyleValue(DocxFontStyle.fontFamily);

            if (string.IsNullOrEmpty(fontFamily))
            {
                string face = node.ExtractAttributeValue("face");

                if (!string.IsNullOrEmpty(face))
                {
                    node.SetExtentedStyle(DocxFontStyle.fontFamily, face);
                }
            }

            string fontSize = node.ExtractOwnStyleValue(DocxFontStyle.fontSize);

            if (string.IsNullOrEmpty(fontSize))
            {
                SetFontSize(node);
            }

            string color = node.ExtractOwnStyleValue(DocxColor.color);

            if(string.IsNullOrEmpty(color))
            {
                color = node.ExtractAttributeValue("color");

                if(!string.IsNullOrEmpty(color))
                {
                    node.SetExtentedStyle(DocxColor.color, color);
                }
            }
        }

        private void Init()
        {
            fontSizes.Add(1, 10);
            fontSizes.Add(2, 13);
            fontSizes.Add(3, 16);
            fontSizes.Add(4, 18);
            fontSizes.Add(5, 24);
            fontSizes.Add(6, 32);
            fontSizes.Add(7, 48);
        }

        internal DocxFont(IOpenXmlContext context)
            : base(context)
        {
            fontSizes = new Dictionary<Int32, Int32>();

            Init();
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "font", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            ApplyStyle(node);

            ProcessElement(node, ref paragraph);
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

            ProcessTextChild(node);
        }
    }
}
