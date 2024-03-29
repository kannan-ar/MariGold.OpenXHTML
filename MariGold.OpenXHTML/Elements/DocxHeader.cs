﻿namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections.Generic;

    internal sealed class DocxHeader : DocxElement, ITextElement
    {
        private Paragraph CreateParagraph(DocxNode node)
        {
            Paragraph para = node.Parent.AppendChild(new Paragraph());
            OnParagraphCreated(node, para);
            return para;
        }

        internal DocxHeader(IOpenXmlContext context) : base(context) { }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "header", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            paragraph = null;
            Paragraph headerParagraph = null;

            ProcessBlockElement(node, ref headerParagraph, properties);
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
