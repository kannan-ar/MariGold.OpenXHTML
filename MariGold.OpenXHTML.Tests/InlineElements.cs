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
    public class InlineElements
    {
        [Test]
        public void FontTagDefault()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<font>test</font>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);
                var text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void SpanDefault()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<span>test</span>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);
                var text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void FontTagWithAttributes()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<font face=\"Arial\" size=\"3\" color=\"red\">test</font>"));

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
                Assert.AreEqual(3, properties.ChildElements.Count);

                RunFonts fonts = properties.ChildElements[0] as RunFonts;
                Assert.IsNotNull(fonts);
                Assert.AreEqual("Arial", fonts.Ascii.Value);

                Word.Color color = properties.ChildElements[1] as Word.Color;
                Assert.AreEqual("FF0000", color.Val.Value);

                FontSize fontSize = properties.ChildElements[2] as FontSize;
                Assert.IsNotNull(fontSize);
                Assert.AreEqual("32", fontSize.Val.Value);

                var text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }
    }
}
