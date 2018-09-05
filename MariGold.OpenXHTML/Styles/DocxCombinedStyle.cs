namespace MariGold.OpenXHTML
{
    using System.Collections.Generic;

    internal static class DocxCombinedStyle
    {
        private static bool MergeTextDecorationStyles(string value, Dictionary<string, string> styles)
        {
            string dictValue;
            string style = string.Empty;
            bool merged = false;

            if (styles.TryGetValue(DocxFontStyle.textDecoration, out dictValue))
            {
                style = DocxFontStyle.textDecoration;
            }
            else if(styles.TryGetValue(DocxFontStyle.textDecorationLine, out dictValue))
            {
                style = DocxFontStyle.textDecorationLine;
            }

            if (!string.IsNullOrEmpty(style) &&
                    !dictValue.CompareStringOrdinalIgnoreCase(value) &&
                    (value.CompareStringOrdinalIgnoreCase(DocxFontStyle.lineThrough) ||
                    value.CompareStringOrdinalIgnoreCase(DocxFontStyle.underLine)))
            {
                styles[style] = string.Concat(dictValue, "|", value);
                merged = true;
            }

            return merged;
        }

        internal static bool MergeGroupStyles(string styleName, string value, Dictionary<string, string> styles)
        {
            bool merged = false;

            switch (styleName)
            {
                case DocxFontStyle.textDecoration:
                case DocxFontStyle.textDecorationLine:
                    merged = MergeTextDecorationStyles(value, styles);
                    break;
            }

            return merged;
        }
    }
}
