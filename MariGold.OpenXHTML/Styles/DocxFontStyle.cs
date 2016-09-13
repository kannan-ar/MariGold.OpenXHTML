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
            styleElement.Append(new RunFonts() { Ascii = CleanFonts(style) });
        }

        internal static void ApplyFontWeight(string style, OpenXmlElement styleElement)
        {
            if (string.Compare(bold, style, StringComparison.InvariantCultureIgnoreCase) == 0 ||
                string.Compare(bolder, style, StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                styleElement.Append(new Bold());
            }
        }

        internal static void ApplyFontStyle(string style, OpenXmlElement styleElement)
        {
            if (string.Compare(italic, style, StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                styleElement.Append(new Italic());
            }
        }

        internal static void ApplyTextDecoration(string style, OpenXmlElement styleElement)
        {
            if (string.Compare(style, underLine, StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                styleElement.Append(new Underline() { Val = UnderlineValues.Single });
            }
            else if (string.Compare(style, lineThrough, StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                styleElement.Append(new Strike());
            }
        }

        internal static void ApplyUnderline(OpenXmlElement styleElement)
        {
            styleElement.Append(new Underline() { Val = UnderlineValues.Single });
        }

        internal static void ApplyFontItalic(OpenXmlElement styleElement)
        {
            styleElement.Append(new Italic());
        }

        internal static void ApplyBold(OpenXmlElement styleElement)
        {
            styleElement.Append(new Bold());
        }

        internal static void ApplyFontSize(string style, OpenXmlElement styleElement)
        {
            decimal fontSize = DocxUnits.HalfPointFromStyle(style);

            if (fontSize != 0)
            {
                fontSize = decimal.Round(fontSize);
                styleElement.Append(new FontSize() { Val = fontSize.ToString("N0") });
            }
        }

        internal static void ApplyFont(int size, bool isBold, OpenXmlElement styleElement)
        {
            FontSize fontSize = new FontSize() { Val = size.ToString() };

            if (isBold)
            {
                styleElement.Append(new Bold());
            }

            styleElement.Append(fontSize);
        }
    }
}
