namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxBr : DocxElement
	{
		internal DocxBr(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "br", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node != null && parent != null)
			{
				if (paragraph == null)
				{
					paragraph = parent.AppendChild(new Paragraph());
					ParagraphCreated(node, paragraph);
				}
				
				Run run = paragraph.AppendChild(new Run(new Break()));
				RunCreated(node, run);
			}
		}
	}
}
