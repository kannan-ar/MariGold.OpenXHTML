namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class InvalidElements
    {
        [Fact]
        public void ScriptElement()
        {
            using MemoryStream mem = new MemoryStream();
            string html = "<script type=\"text/javascript\">console.log('ok');</script>";
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(html));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(0, doc.Document.Body.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void StyleElement()
        {
            using MemoryStream mem = new MemoryStream();
            string html = "<style type=\"text/css\">.cls{font-size:10px;}</style>";
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(html));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(0, doc.Document.Body.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void InvalidHr()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div>1</div><hr><div>2</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(3, doc.Document.Body.ChildElements.Count);

            Word.Paragraph para = doc.Document.Body.ChildElements[0] as Word.Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Word.Run run = para.ChildElements[0] as Word.Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("1", text.InnerText);

            para = doc.Document.Body.ChildElements[2] as Word.Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Word.Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("2", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }
    }
}
