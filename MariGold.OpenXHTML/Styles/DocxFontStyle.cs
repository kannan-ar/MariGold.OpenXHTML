namespace MariGold.OpenXHTML
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal static class DocxFontStyle
    {
        private const int fontMaxLength = 31;

        internal const decimal defaultFontSizeInPixel = 16;
        internal const string fontFamily = "font-family";
        internal const string fontWeight = "font-weight";
        internal const string fontStyle = "font-style";
        internal const string textDecoration = "text-decoration";
        internal const string textDecorationLine = "text-decoration-line";
        internal const string fontSize = "font-size";

        internal const string bold = "bold";
        internal const string bolder = "bolder";
        internal const string italic = "italic";
        internal const string oblique = "oblique";
        internal const string underLine = "underline";
        internal const string lineThrough = "line-through";
        internal const string normal = "normal";

        private static string CleanFonts(string fonts)
        {
            if (string.IsNullOrEmpty(fonts))
            {
                return fonts;
            }

            fonts = fonts.Replace("\"", "").Replace("'", "");

            string[] fontArrary = fonts.Split(new char[] { ',' });
            int totalLength = 0;
            List<string> fontList = new List<string>();

            foreach (string font in fontArrary)
            {
                string newFont = font.Trim();

                if (newFont.Length > 31)
                {
                    continue;
                }

                if ((newFont.Length + totalLength) > fontMaxLength)
                {
                    break;
                }

                fontList.Add(newFont);
                totalLength += newFont.Length;
            }

            return string.Join(",", fontList);
        }

        internal static void ApplyFontFamily(string style, OpenXmlElement styleElement)
        {
            var fontFamilies = CleanFonts(style);
            styleElement.Append(new RunFonts() { Ascii = fontFamilies, ComplexScript = fontFamilies });
        }

        internal static void ApplyFontWeight(string style, OpenXmlElement styleElement)
        {
            if (string.Compare(bold, style, StringComparison.InvariantCultureIgnoreCase) == 0 ||
                string.Compare(bolder, style, StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                styleElement.Append(new Bold());
                styleElement.Append(new BoldComplexScript());
            }
        }

        internal static void ApplyFontStyle(string style, OpenXmlElement styleElement)
        {
            if (string.Compare(italic, style, StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                styleElement.Append(new Italic());
                styleElement.Append(new ItalicComplexScript());
            }
        }

        internal static void ApplyTextDecoration(string style, OpenXmlElement styleElement)
        {
            string[] styles = style.Split(new char[] { '|' });

            foreach (string styleItem in styles)
            {
                if (string.Compare(styleItem, underLine, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    styleElement.Append(new Underline() { Val = UnderlineValues.Single });
                }
                else if (string.Compare(styleItem, lineThrough, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    styleElement.Append(new Strike());
                }
            }
        }

        internal static void ApplyUnderline(OpenXmlElement styleElement)
        {
            styleElement.Append(new Underline() { Val = UnderlineValues.Single });
        }

        internal static void ApplyFontItalic(OpenXmlElement styleElement)
        {
            styleElement.Append(new Italic());
            styleElement.Append(new ItalicComplexScript());
        }

        internal static void ApplyBold(OpenXmlElement styleElement)
        {
            styleElement.Append(new Bold());
            styleElement.Append(new BoldComplexScript());
        }

        internal static void ApplyFontSize(string style, OpenXmlElement styleElement)
        {
            decimal fontSize = DocxUnits.HalfPointFromStyle(style);

            if (fontSize != 0)
            {
                fontSize = decimal.Round(fontSize);
                var fontSizeString = fontSize.ToString("N0");
                styleElement.Append(new FontSize() { Val = fontSizeString });
                styleElement.Append(new FontSizeComplexScript() { Val = fontSizeString });
            }
        }

        internal static void ApplyFont(int size, bool isBold, OpenXmlElement styleElement)
        {
            if (isBold)
            {
                styleElement.Append(new Bold());
                styleElement.Append(new BoldComplexScript());
            }

            var sizeString = size.ToString();
            styleElement.Append(new FontSize() { Val = sizeString });
            styleElement.Append(new FontSizeComplexScript() { Val = sizeString });
        }
    }
}
