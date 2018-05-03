namespace MariGold.OpenXHTML
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxSub : DocxElement, ITextElement
    {
        internal DocxSub(IOpenXmlContext context) : base(context) { }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "sub", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || IsHidden(node))
            {
                return;
            }

            node.SetExtentedStyle(DocxAlignment.verticalAlign, DocxAlignment.sub);

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
