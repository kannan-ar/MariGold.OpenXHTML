namespace MariGold.OpenXHTML
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxBody : DocxElement
    {
        private OpenXmlElement body;

        private void ProcessBody(DocxNode node, ref Paragraph paragraph)
        {
            while (node != null)
            {
                if (node.IsText)
                {
                    if (!IsEmptyText(node, out string text))
                    {
                        if (paragraph == null)
                        {
                            paragraph = body.AppendChild(new Paragraph());
                            OnParagraphCreated(node, paragraph);
                        }

                        Run run = paragraph.AppendChild(new Run(new Text()
                        {
                            Text = ClearHtml(text),
                            Space = SpaceProcessingModeValues.Preserve
                        }));

                        RunCreated(node, run);
                    }
                }
                else
                {
                    node.ParagraphNode = node;
                    node.Parent = body;
                    ProcessChild(node, ref paragraph);
                }

                node = node.Next;
            }
        }

        internal DocxBody(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, "body", StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            body = context.Document.AppendChild(new Body());

            //If the node is body tag, find the first children to process
            if (CanConvert(node))
            {
                if (!node.HasChildren)
                {
                    //Nothing to process. Just return from here.
                    return;
                }

                node = node.Children.ElementAt(0);
            }

            ProcessBody(node, ref paragraph);
        }
    }
}
