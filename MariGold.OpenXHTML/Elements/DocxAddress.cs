namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

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

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || node.Parent == null || IsHidden(node))
            {
                return;
            }

            //Address tag also creats a new block element. Thus clear the existing paragraph
            paragraph = null;
            Paragraph addrParagraph = null;
            node.SetExtentedStyle(DocxFont.fontStyle, DocxFont.italic);

            foreach (DocxNode child in node.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (addrParagraph == null)
                        {
                            addrParagraph = node.Parent.AppendChild(new Paragraph());
                            ParagraphCreated(node.ParagraphNode, addrParagraph);
                        }

                        Run run = addrParagraph.AppendChild(new Run(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        }));

                        RunCreated(node, run);
                    }
                }
                else
                {
                    //Child elements will create on new address paragraph
                    child.ParagraphNode = node.ParagraphNode;
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessChild(child, ref addrParagraph);
                }
            }
        }
    }
}
