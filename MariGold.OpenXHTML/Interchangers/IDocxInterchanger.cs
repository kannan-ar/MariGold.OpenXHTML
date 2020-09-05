namespace MariGold.OpenXHTML
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal interface IDocxInterchanger
    {
        void ProcessImage(IOpenXmlContext context, string imagePath, DocxNode node, ref Paragraph para, Dictionary<string, object> properties);
        void ProcessImage(IOpenXmlContext context, string imagePath, DocxNode node, Dictionary<string, object> properties);
    }
}
