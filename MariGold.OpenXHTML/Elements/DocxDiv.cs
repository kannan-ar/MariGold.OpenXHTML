namespace MariGold.OpenXHTML
{
    using System;
    using System.Collections.Generic;
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

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            //Div creates it's own new paragraph. So old paragraph ends here and creates another one after this div 
            //if there any text!
            paragraph = null;
            Paragraph divParagraph = null;

            ProcessBlockElement(node, ref divParagraph, properties);
        }

        bool ITextElement.CanConvert(DocxNode node)
        {
            return CanConvert(node);
        }

        void ITextElement.Process(DocxNode node, Dictionary<string, object> properties)
        {
            if (IsHidden(node))
            {
                return;
            }

            ProcessTextChild(node, properties);
        }
    }
}
