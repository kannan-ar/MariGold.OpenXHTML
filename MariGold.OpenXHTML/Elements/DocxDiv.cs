namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxDiv : DocxElement, ITextElement
    {
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
                    ProcessParagraph(child, node, ref divParagraph);
                }
                else
                {
                    //ProcessChild forwards the incomming parent to the child element. So any div element inside this div
                    //creates a new paragraph on the parent element.
                    child.ParagraphNode = node;
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessChild(child, ref divParagraph);
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
