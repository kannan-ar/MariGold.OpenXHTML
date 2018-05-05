namespace MariGold.OpenXHTML
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxObject : DocxElement, ITextElement
    {
        internal DocxObject(IOpenXmlContext context) : base(context) { }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "object", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            var interchanger = context.GetInterchanger();
            string data = node.ExtractAttributeValue("data");

            if (data.IsImage())
            {
                interchanger.ProcessImage(context, data, node, ref paragraph);
            }
        }

        bool ITextElement.CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "object", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        void ITextElement.Process(DocxNode node)
        {
            var interchanger = context.GetInterchanger();
            string data = node.ExtractAttributeValue("data");

            if (data.IsImage())
            {
                interchanger.ProcessImage(context, data, node);
            }
        }
    }
}
