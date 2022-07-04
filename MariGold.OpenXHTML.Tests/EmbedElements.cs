namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class EmbedElements
    {
        [Fact]
        public void SimpleObjectWithImage()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.ImagePath = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/objectimage.htm")));

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

        [Fact]
        public void ImageInsideDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");

            doc.ImagePath = path;
            string html = TestUtility.GetHtmlFromFile("/Html/objectimageinsidediv.htm");

            doc.Process(new HtmlParser(html));

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

        [Fact]
        public void ImageParagraphDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.ImagePath = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/objectparagraphimage.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;
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
        public void ImageMultiDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.ImagePath = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/objectimagemultidiv.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(3, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);

            para = doc.Document.Body.ChildElements[2] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test1", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ImageWithBase64Data()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<img src=\"data: image/png;base64, iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" />"));

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

        [Fact]
        public void ImageWithBase64DataWithoutBase64Attr()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<img src=\"data:image/png,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" />"));

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

        [Fact]
        public void ImageWithBase64DataWithoutMediaType()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<img src=\"data:,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" />"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(0, doc.Document.Body.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }
    }
}
