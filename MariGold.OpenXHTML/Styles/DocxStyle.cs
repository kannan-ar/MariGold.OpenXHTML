namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;

    internal static class DocxStyle
    {
        private static IHtmlNode AdjustBackGround(IHtmlNode child, IHtmlNode parent)
        {
            string childBackground;
            string parentBackground;
            string transparent = "transparent";
            string childStyleName = string.Empty;

            if (child.Styles.TryGetValue(DocxColor.backGround, out childBackground))
            {
                childStyleName = DocxColor.backGround;
            }
            else if (child.Styles.TryGetValue(DocxColor.backGroundColor, out childBackground))
            {
                childStyleName = DocxColor.backGroundColor;
            }

            if (string.IsNullOrEmpty(childBackground))
            {
                return child;
            }

            if (!parent.Styles.TryGetValue(DocxColor.backGround, out parentBackground))
            {
                parent.Styles.TryGetValue(DocxColor.backGroundColor, out parentBackground);
            }

            if (string.IsNullOrEmpty(parentBackground))
            {
                return child;
            }

            if (childBackground == transparent && parentBackground != transparent)
            {
                child.Styles[childStyleName] = parentBackground;
            }

            return child;
        }

        internal static IHtmlNode AdjustCSS(IHtmlNode child, IHtmlNode parent)
        {
            IHtmlNode node = child.Clone();

            node = AdjustBackGround(node, parent);

            return node;
        }
    }
}
