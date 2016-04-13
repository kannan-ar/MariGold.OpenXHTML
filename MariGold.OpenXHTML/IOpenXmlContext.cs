namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal interface IOpenXmlContext
	{
		string ImagePath{ get; set; }
		WordprocessingDocument WordprocessingDocument{ get; }
		MainDocumentPart MainDocumentPart{ get; }
		Document Document{ get; }
		
		void Save();
		DocxElement Convert(IHtmlNode node);
		ITextElement ConvertTextElement(IHtmlNode node);
		DocxElement GetBodyElement();
		bool HasNumberingDefinition(NumberFormatValues format);
		void SaveNumberingDefinition(NumberFormatValues format, AbstractNum abstractNum, NumberingInstance numberingInstance);
	}
}
