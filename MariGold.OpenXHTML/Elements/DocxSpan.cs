namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxSpan : WordElement
	{
		public DocxSpan(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool IsBlockLine
		{
			get
			{
				return false;
			}
		}
		
		internal override bool CanConvert(HtmlNode node)
		{
			throw new NotImplementedException();
		}
		
		internal override void Process(HtmlNode node, OpenXmlElement parent)
		{
			throw new NotImplementedException();
		}
	}
}
