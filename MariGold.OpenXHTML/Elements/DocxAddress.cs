﻿namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections.Generic;

    internal sealed class DocxAddress : DocxElement
    {
        internal DocxAddress(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "address", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            //Address tag also creats a new block element. Thus clear the existing paragraph
            paragraph = null;
            Paragraph addrParagraph = null;
            node.SetExtentedStyle(DocxFontStyle.fontStyle, DocxFontStyle.italic);

            ProcessBlockElement(node, ref addrParagraph, properties);
        }
    }
}
