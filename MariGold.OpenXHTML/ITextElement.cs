namespace MariGold.OpenXHTML
{
	internal interface ITextElement
	{
        bool CanConvert(DocxNode node);
        void Process(DocxNode node);
	}
}
