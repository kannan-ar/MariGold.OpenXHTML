namespace MariGold.OpenXHTML.Tests
{
    using System;
    using NUnit.Framework;
    using MariGold.OpenXHTML;
    using System.IO;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Validation;
    using System.Linq;

    [TestFixture]
    public class InvalidElements
    {
        [Test]
        public void ScriptElement()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                string html = "<script type=\"text/javascript\">console.log('ok');</script>";
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser(html));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(0, doc.Document.Body.ChildElements.Count);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void StyleElement()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                string html = "<style type=\"text/css\">.cls{font-size:10px;}</style>";
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser(html));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(0, doc.Document.Body.ChildElements.Count);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }
    }
}
