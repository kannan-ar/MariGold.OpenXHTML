namespace MariGold.OpenXHTML.Tests
{
    using NUnit.Framework;
    using OpenXHTML;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Validation;
    using System.IO;
    using System.Linq;

    [TestFixture]
    public class SimpleHtmlFromFile
    {
        [Test]
        public void EmptyHtmlBody()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/emptybody.htm")));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(0, doc.Document.Body.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void OneSentanceInBody()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/onesentanceinbody.htm")));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual(0, text.ChildElements.Count);
            Assert.AreEqual("This is a test", text.InnerText.Trim());

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void OnePTag()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/oneptag.htm")));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual(0, text.ChildElements.Count);
            Assert.AreEqual("Test", text.InnerText.Trim());

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void PTagWithStyle()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/ptagwithstyle.htm")));

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
            Assert.AreEqual("Arial,Verdana", fonts.Ascii.Value);

            Bold bold = properties.ChildElements[1] as Bold;
            Assert.IsNotNull(bold);

            FontSize fontSize = properties.ChildElements[2] as FontSize;
            Assert.IsNotNull(fontSize);
            Assert.AreEqual("24", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual(0, text.ChildElements.Count);
            Assert.AreEqual("Test", text.InnerText.Trim());


            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void ImageInsideATag()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.ImagePath = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/imageinsideatag.htm")));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual(1, link.ChildElements.Count);

            Run run = link.ChildElements[0] as Run;
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
        public void ImageInsideATagWithBaseURL()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.BaseURL = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/imageinsideatag.htm")));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(1, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual(1, link.ChildElements.Count);

            Run run = link.ChildElements[0] as Run;
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
        public void BootstrapCDN()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            doc.BaseURL = "https://maxcdn.bootstrapcdn.com";
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/relativestylesheet.htm")));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(2, run.ChildElements.Count);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual(0, text.ChildElements.Count);
            Assert.AreEqual("test", text.InnerText.Trim());

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void ImageWithSpace()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.ImagePath = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/imagewithspace.htm")));

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
    }
}
