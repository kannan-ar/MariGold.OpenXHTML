namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxSpan : DocxElement, ITextElement
    {
        internal DocxSpan(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "span", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            ProcessElement(node, ref paragraph);
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
