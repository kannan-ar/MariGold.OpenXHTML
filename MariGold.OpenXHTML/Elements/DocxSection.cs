namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections.Generic;

    internal sealed class DocxSection : DocxElement
    {
        private const string tag = "section";

        internal DocxSection(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, tag, StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            paragraph = null;
            Paragraph sectionParagraph = null;

            ProcessBlockElement(node, ref sectionParagraph, properties);
        }
    }
}
