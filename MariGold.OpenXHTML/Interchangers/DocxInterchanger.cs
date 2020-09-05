namespace MariGold.OpenXHTML
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal class DocxInterchanger : IDocxInterchanger
    {
        private DocxNode GetImageNode(string imagePath)
        {
            var attributes = new Dictionary<string, string>
            {
                { "src", imagePath }
            };

            return new DocxNode(new DocxHtmlNode(attributes));
        }

        public void ProcessImage(IOpenXmlContext context, string imagePath, DocxNode node, ref Paragraph para, Dictionary<string, object> properties)
        {
            DocxImage image = new DocxImage(context);
            DocxNode docxNode = GetImageNode(imagePath);
            docxNode.Parent = node.Parent;
            image.Process(docxNode, ref para, properties);
        }

        public void ProcessImage(IOpenXmlContext context, string imagePath, DocxNode node, Dictionary<string, object> properties)
        {
            ITextElement image = new DocxImage(context);
            DocxNode docxNode = GetImageNode(imagePath);
            docxNode.Parent = node.Parent;

            image.Process(docxNode, properties);
        }
    }
}
