namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class SimpleHtmlFromFile
    {
        [Fact]
        public void EmptyHtmlBody()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/emptybody.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(0, doc.Document.Body.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void OneSentanceInBody()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/onesentanceinbody.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("This is a test", text.InnerText.Trim());

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void OnePTag()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/oneptag.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("Test", text.InnerText.Trim());

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void PTagWithStyle()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/ptagwithstyle.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            Assert.Equal(3, properties.ChildElements.Count);

            RunFonts fonts = properties.ChildElements[0] as RunFonts;
            Assert.NotNull(fonts);
            Assert.Equal("Arial,Verdana", fonts.Ascii.Value);

            Bold bold = properties.ChildElements[1] as Bold;
            Assert.NotNull(bold);

            FontSize fontSize = properties.ChildElements[2] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("24", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("Test", text.InnerText.Trim());


            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ImageInsideATag()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.ImagePath = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/imageinsideatag.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;
            Assert.NotNull(link);
            Assert.Equal(1, link.ChildElements.Count);

            Run run = link.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ImageInsideATagWithBaseURL()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.BaseURL = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/imageinsideatag.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;
            Assert.NotNull(link);
            Assert.Equal(1, link.ChildElements.Count);

            Run run = link.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void BootstrapCDN()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            doc.BaseURL = "https://maxcdn.bootstrapcdn.com";
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/relativestylesheet.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            Run run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText.Trim());

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ImageWithSpace()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.ImagePath = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/imagewithspace.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }
    }
}
