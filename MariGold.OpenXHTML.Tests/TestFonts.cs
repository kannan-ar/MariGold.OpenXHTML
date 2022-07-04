namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class TestFonts
    {
        [Fact]
        public void DivFontBold()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='font-weight:bold'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            OpenXmlElement para = doc.Document.Body.ChildElements[0];

            Assert.True(para is Paragraph);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Assert.NotNull(run.RunProperties);
            Assert.Equal(1, run.RunProperties.ChildElements.Count);
            Bold bold = run.RunProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivFontFamily()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='font-family:arial'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            OpenXmlElement para = doc.Document.Body.ChildElements[0];

            Assert.True(para is Paragraph);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Assert.NotNull(run.RunProperties);
            Assert.Equal(1, run.RunProperties.ChildElements.Count);
            RunFonts fonts = run.RunProperties.ChildElements[0] as RunFonts;
            Assert.NotNull(fonts);
            Assert.Equal("arial", fonts.Ascii.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivMultipleFontFamily()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='font-family:Arial, Georgia, Serif'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            OpenXmlElement para = doc.Document.Body.ChildElements[0];

            Assert.True(para is Paragraph);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Assert.NotNull(run.RunProperties);
            Assert.Equal(1, run.RunProperties.ChildElements.Count);
            RunFonts fonts = run.RunProperties.ChildElements[0] as RunFonts;
            Assert.NotNull(fonts);
            Assert.Equal("Arial,Georgia,Serif", fonts.Ascii.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void ATagWithFont()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com' style='font-family:arial'>test</a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            OpenXmlElement para = doc.Document.Body.ChildElements[0];

            Assert.True(para is Paragraph);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;
            Assert.NotNull(link);

            Run run = link.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Assert.NotNull(run.RunProperties);
            Assert.Equal(2, run.RunProperties.ChildElements.Count);

            RunStyle runStyle = run.RunProperties.ChildElements[0] as RunStyle;
            Assert.NotNull(runStyle);

            RunFonts fonts = run.RunProperties.ChildElements[1] as RunFonts;
            Assert.NotNull(fonts);
            Assert.Equal("arial", fonts.Ascii.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }
    }
}
