namespace MariGold.OpenXHTML
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxUnderline : DocxElement, ITextElement
    {
        internal DocxUnderline(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "u", StringComparison.InvariantCultureIgnoreCase) == 0 ||
                string.Compare(node.Tag, "ins", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || IsHidden(node))
            {
                return;
            }
            
            node.SetExtentedStyle(DocxFontStyle.textDecoration, DocxFontStyle.underLine);

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

            node.SetExtentedStyle(DocxFontStyle.textDecoration, DocxFontStyle.underLine);
            ProcessTextChild(node);
        }
    }
}
