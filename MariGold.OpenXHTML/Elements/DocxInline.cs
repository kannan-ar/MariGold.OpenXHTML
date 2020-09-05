namespace MariGold.OpenXHTML
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxInline : DocxElement, ITextElement
    {
        private readonly string[] nonTextTags = { "script", "style" };

        private bool IsTextTag(string tag)
        {
            foreach (string nonTextTag in nonTextTags)
            {
                if (string.Compare(tag, nonTextTag, true) == 0)
                {
                    return false;
                }
            }

            return true;
        }

        internal DocxInline(IOpenXmlContext context) : base(context) { }

        internal override bool CanConvert(DocxNode node)
        {
            return IsTextTag(node.Tag);
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            ProcessElement(node, ref paragraph, properties);
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
