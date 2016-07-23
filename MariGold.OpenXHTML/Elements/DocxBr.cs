namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxBr : DocxElement, ITextElement
	{
		internal DocxBr(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(DocxNode node)
		{
			return string.Compare(node.Tag, "br", StringComparison.InvariantCultureIgnoreCase) == 0;
		}

        internal override void Process(DocxNode node, ref Paragraph paragraph)
		{
            if (!node.IsNull() && node.Parent != null || IsHidden(node))
			{
				if (paragraph == null)
				{
                    paragraph = node.Parent.AppendChild(new Paragraph());
                    ParagraphCreated(node.ParagraphNode, paragraph);
				}
				
				Run run = paragraph.AppendChild(new Run(new Break()));
                RunCreated(node, run);
			}
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

            node.Parent.AppendChild(new Run(new Break()));
        }
	}
}
