namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml;
    using MariGold.HtmlParser;
    using System;
    using System.Collections.Generic;

    internal sealed class DocxNode
    {
        private readonly IHtmlNode node;

        private DocxNode paragraphNode;
        private OpenXmlElement parent;
        private Dictionary<string, string> extentedStyles;
        private readonly Dictionary<string, string> styles;
        private readonly Dictionary<string, string> inheritedStyles;

        private void SetExtentedStyles(Dictionary<string, string> extentedStyles)
        {
            this.extentedStyles = new Dictionary<string, string>();

            foreach (var style in extentedStyles)
            {
                if (!node.Styles.ContainsKey(style.Key))
                {
                    this.extentedStyles.Add(style.Key, style.Value);
                }
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
            this.node = node ?? throw new ArgumentNullException("node");
            this.styles = node.Styles;
            this.inheritedStyles = node.InheritedStyles;
            this.extentedStyles = new Dictionary<string, string>();
        }

        internal bool IsNull()
        {
            return node == null;
        }

        internal string ExtractAttributeValue(string attributeName)
        {
            foreach (KeyValuePair<string, string> attribute in node.Attributes)
            {
                if (string.Equals(attributeName, attribute.Key, StringComparison.OrdinalIgnoreCase))
                {
                    return attribute.Value;
                }
            }

            return string.Empty;
        }

        internal string ExtractStyleValue(string styleName)
        {
            foreach (KeyValuePair<string, string> style in styles)
            {
                if (string.Equals(styleName, style.Key, StringComparison.OrdinalIgnoreCase))
                {
                    return style.Value;
                }
            }

            foreach (KeyValuePair<string, string> style in extentedStyles)
            {
                if (string.Equals(styleName, style.Key, StringComparison.OrdinalIgnoreCase))
                {
                    return style.Value;
                }
            }

            foreach (KeyValuePair<string, string> style in inheritedStyles)
            {
                if (string.Equals(styleName, style.Key, StringComparison.OrdinalIgnoreCase))
                {
                    return style.Value;
                }
            }

            return string.Empty;
        }

        internal string ExtractOwnStyleValue(string styleName)
        {
            foreach (KeyValuePair<string, string> style in styles)
            {
                if (string.Equals(styleName, style.Key, StringComparison.OrdinalIgnoreCase))
                {
                    return style.Value;
                }
            }

            return string.Empty;
        }

        internal string ExtractInheritedStyleValue(string styleName)
        {
            foreach (KeyValuePair<string, string> style in extentedStyles)
            {
                if (string.Equals(styleName, style.Key, StringComparison.OrdinalIgnoreCase))
                {
                    return style.Value;
                }
            }

            foreach (KeyValuePair<string, string> style in inheritedStyles)
            {
                if (string.Equals(styleName, style.Key, StringComparison.OrdinalIgnoreCase))
                {
                    return style.Value;
                }
            }

            return string.Empty;
        }

        internal void SetExtentedStyle(string styleName, string value)
        {
            if (!DocxCombinedStyle.MergeGroupStyles(styleName, value, this.extentedStyles))
            {
                this.extentedStyles[styleName] = value;
            }
        }

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

            foreach (string styleName in styleNames)
            {
                inheritedStyles.Remove(styleName);
            }
        }
    }
}
