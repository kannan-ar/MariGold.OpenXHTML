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
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "br", StringComparison.InvariantCultureIgnoreCase) == 0;
		}

        internal override void Process(DocxProperties properties, ref Paragraph paragraph)
		{
            if (properties.CurrentNode != null && properties.Parent != null)
			{
				if (paragraph == null)
				{
                    paragraph = properties.Parent.AppendChild(new Paragraph());
                    ParagraphCreated(properties.ParagraphNode, paragraph);
				}
				
				Run run = paragraph.AppendChild(new Run(new Break()));
                RunCreated(properties.CurrentNode, run);
			}
		}

        bool ITextElement.CanConvert(IHtmlNode node)
        {
            return CanConvert(node);
        }

        void ITextElement.Process(DocxProperties properties)
        {
            properties.Parent.AppendChild(new Run(new Break()));
        }
	}
}
