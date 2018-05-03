namespace MariGold.OpenXHTML.Elements
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxObject : DocxElement, ITextElement
    {
        internal DocxObject(IOpenXmlContext context) : base(context) { }

        internal override bool CanConvert(DocxNode node)
        {
            throw new NotImplementedException();
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            throw new NotImplementedException();
        }

        bool ITextElement.CanConvert(DocxNode node)
        {
            throw new NotImplementedException();
        }

        void ITextElement.Process(DocxNode node)
        {
            throw new NotImplementedException();
        }
    }
}
