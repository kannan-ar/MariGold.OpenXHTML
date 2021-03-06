﻿namespace MariGold.OpenXHTML.Tests
{
    using NUnit.Framework;
    using OpenXHTML;
    using System.IO;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Validation;
    using System.Linq;

    [TestFixture]
    public class EmbedElements
    {
        [Test]
        public void SimpleObjectWithImage()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.ImagePath = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/objectimage.htm")));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.IsNotNull(image);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void ImageInsideDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");

            doc.ImagePath = path;
            string html = TestUtility.GetHtmlFromFile("/Html/objectimageinsidediv.htm");

            doc.Process(new HtmlParser(html));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.IsNotNull(image);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void ImageParagraphDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.ImagePath = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/objectparagraphimage.htm")));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("test", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.IsNotNull(image);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void ImageMultiDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.ImagePath = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/objectimagemultidiv.htm")));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(3, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("test", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.IsNotNull(image);

            para = doc.Document.Body.ChildElements[2] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("test1", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void ImageWithBase64Data()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<img src=\"data: image/png;base64, iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" />"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.IsNotNull(image);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void ImageWithBase64DataWithoutBase64Attr()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<img src=\"data:image/png,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" />"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.IsNotNull(image);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void ImageWithBase64DataWithoutMediaType()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<img src=\"data:,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" />"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(0, doc.Document.Body.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }
    }
}
