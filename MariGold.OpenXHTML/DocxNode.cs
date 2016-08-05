namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;

    internal sealed class DocxNode
    {
        private readonly IHtmlNode node;

        private DocxNode paragraphNode;
        private OpenXmlElement parent;
        private Dictionary<string, string> extentedStyles;
        private Dictionary<string, string> styles;

        private void SetExtentedStyles(Dictionary<string, string> extentedStyles)
        {
            this.extentedStyles = new Dictionary<string,string>();

            foreach(var style in extentedStyles)
            {
                this.extentedStyles.Add(style.Key, style.Value);
            }
        }

        internal string Tag
        {
            get
            {
                return node.Tag;
            }
        }

        internal string Html
        {
            get
            {
                return node.Html;
            }
        }

        internal string InnerHtml
        {
            get
            {
                return node.InnerHtml;
            }
        }

        internal bool IsText
        {
            get
            {
                return node.IsText;
            }
        }

        internal DocxNode Next
        {
            get
            {
                if (node.Next == null)
                {
                    return null;
                }

                return new DocxNode(node.Next);
            }
        }

        internal DocxNode Previous
        {
            get
            {
                if (node.Previous == null)
                {
                    return null;
                }

                return new DocxNode(node.Previous);
            }
        }

        internal bool HasChildren
        {
            get
            {
                return node.HasChildren;
            }
        }

        internal IEnumerable<DocxNode> Children
        {
            get
            {
                foreach (var child in node.Children)
                {
                    yield return new DocxNode(child);
                }
            }
        }
        /*
        internal IHtmlNode CurrentNode
        {
            get
            {
                return node;
            }
        }
        */
        internal DocxNode ParagraphNode
        {
            get
            {
                return paragraphNode ?? this;
            }

            set
            {
                paragraphNode = value;
            }
        }

        internal Dictionary<string, string> Styles
        {
            get
            {
                return styles;
            }
        }

        internal OpenXmlElement Parent
        {
            get
            {
                return parent;
            }

            set
            {
                parent = value;
            }
        }

        internal DocxNode(IHtmlNode node)
        {
            if (node == null)
            {
                throw new ArgumentNullException("node");
            }

            this.node = node;
            this.styles = node.Styles;
            this.extentedStyles = new Dictionary<string, string>();
        }

        /*
        internal DocxNode(IHtmlNode currentNode, OpenXmlElement parent)
        {
            this.node = currentNode;
            this.parent = parent;

            Init();
        }

        internal DocxNode(IHtmlNode currentNode, IHtmlNode paragraphNode, OpenXmlElement parent)
        {
            this.node = currentNode;
            this.paragraphNode = paragraphNode;
            this.parent = parent;

            Init();
        }
        */

        internal bool IsNull()
        {
            return node == null;
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
            foreach (KeyValuePair<string, string> style in extentedStyles)
            {
                if (string.Compare(styleName, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    return style.Value;
                }
            }

            foreach (KeyValuePair<string, string> style in styles)
            {
                if (string.Compare(styleName, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    return style.Value;
                }
            }

            return string.Empty;
        }

        internal void SetExtentedStyle(string styleName, string value)
        {
            this.extentedStyles[styleName] = value;
        }

        /*
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
        */
        /*
        internal void CopyStyles(DocxNode toNode, params string[] styles)
        {
            if (toNode.IsNull())
            {
                return;
            }

            foreach (string style in styles)
            {
                string value;

                if (this.Styles.TryGetValue(style, out value) && !toNode.Styles.ContainsKey(style))
                {
                    toNode.Styles.Add(style, value);
                }
            }
        }
        */
        internal void CopyExtentedStyles(DocxNode toNode)
        {
            if (toNode.IsNull())
            {
                return;
            }

            toNode.SetExtentedStyles(this.extentedStyles);
        }

        internal void RemoveStyles(params string[] styleNames)
        {
            foreach (string styleName in styleNames)
            {
                styles.Remove(styleName);
            }
        }
    }
}
