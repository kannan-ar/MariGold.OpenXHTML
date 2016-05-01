namespace MariGold.OpenXHTML.Tests
{
    using System;
    using NUnit.Framework;
    using MariGold.OpenXHTML;
    using System.IO;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Validation;
    using System.Linq;

    [TestFixture]
    public class BlockElements
    {
        [Test]
        public void HeaderElement()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<header>test</header>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(paragraph);

                Run run = paragraph.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual(0, text.ChildElements.Count);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void FooterElement()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<footer>test</footer>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(paragraph);

                Run run = paragraph.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual(0, text.ChildElements.Count);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }
    }
}
