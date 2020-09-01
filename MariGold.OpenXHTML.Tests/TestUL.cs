namespace MariGold.OpenXHTML.Tests
{
    using System.IO;
    using System.Linq;
    using NUnit.Framework;
    using OpenXHTML;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Validation;

    [TestFixture]
    public class TestUL
    {
        [Test]
        public void TestOne()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ul><li>One</li></ul>"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("One", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.AreEqual(0, errors.Count());
        }
    }
}
