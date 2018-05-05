namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;

    internal interface IDocxInterchanger
    {
        void ProcessImage(IOpenXmlContext context, string imagePath, DocxNode node, ref Paragraph para);
        void ProcessImage(IOpenXmlContext context, string imagePath, DocxNode node);
    }
}
