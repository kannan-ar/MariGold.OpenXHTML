namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxDiv : DocxElement, ITextElement
    {
        private Paragraph CreateParagraph(DocxNode node)
        {
            Paragraph para = node.Parent.AppendChild(new Paragraph());
            ParagraphCreated(node, para);
            return para;
        }

        internal DocxDiv(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "div", StringComparison.InvariantCultureIgnoreCase) == 0 ||
            string.Compare(node.Tag, "p", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            //Div creates it's own new paragraph. So old paragraph ends here and creats another one after this div 
            //if there any text!
            paragraph = null;
            Paragraph divParagraph = null;

            foreach (DocxNode child in node.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (divParagraph == null)
                        {
                            divParagraph = CreateParagraph(node);
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
                    child.ParagraphNode = node;
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessChild(child, ref divParagraph);
                    //ProcessChild(new DocxNode(DocxStyle.AdjustCSS(child, node), node, node.Parent), ref divParagraph);
                }
            }
        }

        bool ITextElement.CanConvert(DocxNode node)
        {
            return CanConvert(node);
        }

        void ITextElement.Process(DocxNode node)
        {
            if (IsHidden(node))
            {
                return;
            }

            ProcessTextChild(node);
        }
    }
}
