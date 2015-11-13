namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal interface IWordContext
	{
		WordprocessingDocument WordprocessingDocument{ get; }
		MainDocumentPart MainDocumentPart{ get; }
		Document Document{ get; }
		
		void Clear();
		
		WordElement Convert(HtmlNode node);
	}
}
