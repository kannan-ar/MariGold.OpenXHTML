namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal interface IOpenXmlContext
	{
		string ImagePath{ get; set; }
		string BaseURL{ get; set; }
        string UriSchema { get; set; }
		WordprocessingDocument WordprocessingDocument{ get; }
		MainDocumentPart MainDocumentPart{ get; }
		Document Document{ get; }
        IParser Parser { get; }

		void Save();
		DocxElement Convert(DocxNode node);
        ITextElement ConvertTextElement(DocxNode node);
		DocxElement GetBodyElement();
		bool HasNumberingDefinition(NumberFormatValues format);
		void SaveNumberingDefinition(NumberFormatValues format, AbstractNum abstractNum, NumberingInstance numberingInstance);
        void SetParser(IParser parser);
	}
}
