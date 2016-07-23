namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	
	internal interface ITextElement
	{
        bool CanConvert(DocxNode node);
        void Process(DocxNode node);
	}
}
