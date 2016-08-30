namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxInline : DocxElement, ITextElement
    {
        public DocxInline(IOpenXmlContext context) : base(context) { }

        internal override bool CanConvert(DocxNode node)
        {
            return true;//It can convert anything!!!
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
