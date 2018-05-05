namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using System.Collections.Generic;

    internal sealed class DocxHtmlNode : IHtmlNode
    {
        private readonly Dictionary<string, string> attributes;
        private readonly Dictionary<string, string> styles;
        private readonly Dictionary<string, string> inheritedStyles;

        public Dictionary<string, string> Attributes
        {
            get
            {
                return attributes;
            }
        }

        public IEnumerable<IHtmlNode> Children { get; }
        public bool HasChildren { get; }
        public string Html { get; }
        public Dictionary<string, string> InheritedStyles
        {
            get
            {
                return inheritedStyles;
            }
        }

        public string InnerHtml { get; }
        public bool IsText { get; }
        public IHtmlNode Next { get; }
        public IHtmlNode Parent { get; }
        public IHtmlNode Previous { get; }
        public bool SelfClosing { get; }
        public Dictionary<string, string> Styles
        {
            get
            {
                return styles;
            }
        }

        public string Tag { get; }

        internal DocxHtmlNode(Dictionary<string, string> attributes)
        {
            this.attributes = attributes;
            this.styles = new Dictionary<string, string>();
            this.inheritedStyles = new Dictionary<string, string>();
        }

        public IHtmlNode Clone()
        {
            throw new NotImplementedException();
        }
    }
}
