namespace MariGold.OpenXHTML
{
    using System;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal interface IOpenXmlContext
	{
		string ImagePath{ get; set; }
		string BaseURL{ get; set; }
        string UriSchema { get; set; }
		WordprocessingDocument WordprocessingDocument{ get; }
		MainDocumentPart MainDocumentPart{ get; }
		Document Document{ get; }
        IParser Parser { get; }
        Int16 ListNumberId { get; set; }
        Int32 RelationshipId { get; set; }

        void Save();
		DocxElement Convert(DocxNode node);
        ITextElement ConvertTextElement(DocxNode node);
		DocxElement GetBodyElement();
		void SaveNumberingDefinition(Int16 numberId, AbstractNum abstractNum, NumberingInstance numberingInstance);
        void SetParser(IParser parser);
        IDocxInterchanger GetInterchanger();
	}
}
