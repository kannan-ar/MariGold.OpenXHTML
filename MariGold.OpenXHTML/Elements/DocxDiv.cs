namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxDiv : DocxElement, ITextElement
    {
        private Paragraph CreateParagraph(DocxProperties properties)
        {
            Paragraph para = properties.Parent.AppendChild(new Paragraph());
            ParagraphCreated(properties.CurrentNode, para);
            return para;
        }

        internal DocxDiv(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(IHtmlNode node)
        {
            return string.Compare(node.Tag, "div", StringComparison.InvariantCultureIgnoreCase) == 0 ||
            string.Compare(node.Tag, "p", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxProperties properties, ref Paragraph paragraph)
        {
            if (properties.CurrentNode == null || properties.Parent == null)
            {
                return;
            }

            //Div creates it's own new paragraph. So old paragraph ends here and creats another one after this div 
            //if there any text!
            paragraph = null;
            Paragraph divParagraph = null;

            foreach (IHtmlNode child in properties.CurrentNode.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (divParagraph == null)
                        {
                            divParagraph = CreateParagraph(properties);
                        }

                        Run run = divParagraph.AppendChild(new Run(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        }));

                        RunCreated(child, run);
                    }
                }
                else
                {
                    //ProcessChild forwards the incomming parent to the child element. So any div element inside this div
                    //creates a new paragraph on the parent element.
                    ProcessChild(new DocxProperties(child, properties.CurrentNode, properties.Parent), ref divParagraph);
                }
            }
        }

        bool ITextElement.CanConvert(IHtmlNode node)
        {
            return CanConvert(node);
        }

        void ITextElement.Process(DocxProperties properties)
        {
            ProcessTextChild(properties);
        }
    }
}
