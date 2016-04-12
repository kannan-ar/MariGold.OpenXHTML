namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	
	internal interface ITextElement
	{
		bool CanConvert(IHtmlNode node);
		void Process(IHtmlNode node, OpenXmlElement parent);
	}
}
