namespace MariGold.OpenXHTML.Tests
{
    using NUnit.Framework;
    using OpenXHTML;
    using System.IO;
    using DocumentFormat.OpenXml.Validation;
    using System.Linq;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

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

        [Test]
        public void InvalidHr()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<div>1</div><hr><div>2</div>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(3, doc.Document.Body.ChildElements.Count);

                Word.Paragraph para = doc.Document.Body.ChildElements[0] as Word.Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Word.Run run = para.ChildElements[0] as Word.Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual("1", text.InnerText);

                para = doc.Document.Body.ChildElements[2] as Word.Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                run = para.ChildElements[0] as Word.Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual("2", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }
    }
}
