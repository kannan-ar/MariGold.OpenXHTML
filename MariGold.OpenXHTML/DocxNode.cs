namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using System.Collections.Generic;

    internal sealed class DocxNode
    {
        private readonly IHtmlNode node;

        internal string Tag
        {
            get
            {
                return node.Tag;
            }
        }

        internal DocxNode(IHtmlNode node)
        {
            if (node == null)
            {
                throw new ArgumentNullException("node");
            }

            this.node = node;
        }

        internal string ExtractAttributeValue(string attributeName)
        {
            foreach (KeyValuePair<string, string> attribute in node.Attributes)
            {
                if (string.Compare(attributeName, attribute.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    return attribute.Value;
                }
            }

            return string.Empty;
        }

        internal string ExtractStyleValue(string styleName)
        {
            foreach (KeyValuePair<string, string> style in node.Styles)
            {
                if (string.Compare(styleName, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    return style.Value;
                }
            }

            return string.Empty;
        }

        internal void SetStyleValue(string styleName, string value)
        {
            string key = string.Empty;

            foreach (KeyValuePair<string, string> style in node.Styles)
            {
                if (string.Compare(styleName, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    key = style.Key;
                }
            }

            if (string.IsNullOrEmpty(key))
            {
                key = styleName;
            }

            Dictionary<string, string> styles = node.Styles;
            styles[key] = value;
            node.Styles = styles;
        }

        internal void SetStyleValues(Dictionary<string, string> newStyles)
        {
            Dictionary<string, string> styles = node.Styles;

            foreach (KeyValuePair<string, string> newStyle in newStyles)
            {
                string styleName = newStyle.Key;

                foreach (string key in styles.Keys)
                {
                    if (string.Compare(key, newStyle.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        styleName = key;
                        break;
                    }
                }

                styles[styleName] = newStyle.Value;
            }

            node.Styles = styles;
        }

        internal void CopyStyles(IHtmlNode toNode, params string[] styles)
        {
            if (toNode == null)
            {
                return;
            }

            foreach (string style in styles)
            {
                string value;

                if (node.Styles.TryGetValue(style, out value) && !toNode.Styles.ContainsKey(style))
                {
                    toNode.Styles.Add(style, value);
                }
            }
        }
    }
}
