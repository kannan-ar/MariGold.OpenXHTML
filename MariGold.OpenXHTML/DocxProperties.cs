namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;

    internal sealed class DocxProperties
    {
        private readonly IHtmlNode currentNode;
        private readonly IHtmlNode paragraphNode;
        private readonly OpenXmlElement parent;

        internal DocxProperties(IHtmlNode currentNode, OpenXmlElement parent)
        {
            this.currentNode = currentNode;
            this.parent = parent;
        }

        internal DocxProperties(IHtmlNode currentNode, IHtmlNode paragraphNode, OpenXmlElement parent)
        {
            this.currentNode = currentNode;
            this.paragraphNode = paragraphNode;
            this.parent = parent;
        }

        internal IHtmlNode CurrentNode
        {
            get
            {
                return currentNode;
            }
        }

        internal IHtmlNode ParagraphNode
        {
            get
            {
                return paragraphNode ?? currentNode;
            }
        }

        internal OpenXmlElement Parent
        {
            get
            {
                return parent;
            }
        }
    }
}
