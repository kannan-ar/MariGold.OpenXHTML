namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.Collections.Generic;

    internal interface IDocxInterchanger
    {
        void ProcessImage(IOpenXmlContext context, string imagePath, DocxNode node, ref Paragraph para, Dictionary<string, object> properties);
        void ProcessImage(IOpenXmlContext context, string imagePath, DocxNode node, Dictionary<string, object> properties);
    }
}
