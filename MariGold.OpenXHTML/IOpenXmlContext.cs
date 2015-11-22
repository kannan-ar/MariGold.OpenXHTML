namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal interface IOpenXmlContext
	{
		WordprocessingDocument WordprocessingDocument{ get; }
		MainDocumentPart MainDocumentPart{ get; }
		Document Document{ get; }
		Paragraph LastParagraph{ get; set; }
		
		void Clear();
		DocxElement Convert(HtmlNode node);
		DocxElement GetBodyElement();
	}
}
