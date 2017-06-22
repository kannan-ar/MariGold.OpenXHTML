namespace MariGold.OpenXHTML
{
    using System.Collections.Generic;

    internal static class DocxCombinedStyle
    {
        private static bool MergeTextDecorationStyles(string value, Dictionary<string, string> styles)
        {
            string dictValue;
            bool merged = false;

            if (styles.TryGetValue(DocxFontStyle.textDecoration, out dictValue))
            {
                if (!dictValue.CompareStringInvariantCultureIgnoreCase(value) &&
                    (value.CompareStringInvariantCultureIgnoreCase(DocxFontStyle.lineThrough) ||
                    value.CompareStringInvariantCultureIgnoreCase(DocxFontStyle.underLine)))
                {
                    styles[DocxFontStyle.textDecoration] = string.Concat(dictValue, "|", value);
                    merged = true;
                }
            }

            return merged;
        }

        internal static bool MergeGroupStyles(string styleName, string value, Dictionary<string, string> styles)
        {
            bool merged = false;

            switch (styleName)
            {
                case DocxFontStyle.textDecoration:
                    merged = MergeTextDecorationStyles(value, styles);
                    break;
            }

            return merged;
        }
    }
}
