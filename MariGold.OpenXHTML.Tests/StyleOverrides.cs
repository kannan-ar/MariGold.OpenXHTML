namespace MariGold.OpenXHTML.Tests
{
    using NUnit.Framework;
    using OpenXHTML;
    using DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Validation;
    using System.IO;
    using System.Linq;

    [TestFixture]
    public class StyleOverrides
    {
        [Test]
        public void OwnStyleOverride()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/ownstyleoverride.htm")));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);

            Bold bold = properties.ChildElements[0] as Bold;
            Assert.IsNotNull(bold);

            FontSize fontSize = properties.ChildElements[1] as FontSize;
            Assert.IsNotNull(fontSize);
            Assert.AreEqual("46", fontSize.Val.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }
    }
}
