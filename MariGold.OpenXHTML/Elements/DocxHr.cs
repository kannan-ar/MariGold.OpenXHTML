namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxHr : DocxElement
	{
		internal DocxHr(IOpenXmlContext context)
			: base(context)
		{
			
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "hr", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node != null && parent != null)
			{
				paragraph = null;
				
				Paragraph hrParagraph = parent.AppendChild(new Paragraph());
				ParagraphCreated(node, hrParagraph);
				
				if (hrParagraph.ParagraphProperties == null)
				{
					hrParagraph.ParagraphProperties = new ParagraphProperties();
				}
				
				ParagraphBorders paragraphBorders = new ParagraphBorders();
				DocxBorder.ApplyDefaultBorder<TopBorder>(paragraphBorders);
				hrParagraph.ParagraphProperties.Append(paragraphBorders);
				
				Run run = hrParagraph.AppendChild(new Run(new Text()));
				RunCreated(node, run);
				
			}
		}
	}
}
