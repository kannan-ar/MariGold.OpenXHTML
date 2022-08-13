namespace MariGold.OpenXHTML
{
    using System.Collections.Generic;

    internal interface ITextElement
    {
        bool CanConvert(DocxNode node);
        void Process(DocxNode node, Dictionary<string, object> properties);
    }
}
