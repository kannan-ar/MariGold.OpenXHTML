namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class InlineElements
    {
        [Fact]
        public void FontTagDefault()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<font>test</font>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            var text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SpanDefault()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<span>test</span>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            var text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void FontTagWithAttributes()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<font face=\"Arial\" size=\"3\" color=\"red\">test</font>"));

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
            Assert.Equal("Arial", fonts.Ascii.Value);

            Word.Color color = properties.ChildElements[1] as Word.Color;
            Assert.Equal("FF0000", color.Val.Value);

            FontSize fontSize = properties.ChildElements[2] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("32", fontSize.Val.Value);

            var text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TestElement()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<test>one</test>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            var text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("one", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void QuoteElement()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<q>one</q>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(3, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            var text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("\"", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("one", text.InnerText);

            run = para.ChildElements[2] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("\"", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SpanQuoteElement()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<span>test</span><q>one</q>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(4, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            var text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("\"", text.InnerText);

            run = para.ChildElements[2] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("one", text.InnerText);

            run = para.ChildElements[3] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("\"", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void Superscript()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<sup>test</sup>"));

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
            Assert.Equal(1, properties.ChildElements.Count);

            VerticalTextAlignment verticalTextAlignment = properties.ChildElements[0] as VerticalTextAlignment;
            Assert.NotNull(verticalTextAlignment);
            Assert.Equal(VerticalPositionValues.Superscript, verticalTextAlignment.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void Subscript()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<sub>test</sub>"));

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
            Assert.Equal(1, properties.ChildElements.Count);

            VerticalTextAlignment verticalTextAlignment = properties.ChildElements[0] as VerticalTextAlignment;
            Assert.NotNull(verticalTextAlignment);
            Assert.Equal(VerticalPositionValues.Subscript, verticalTextAlignment.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void StrikeDel()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<del>test</del>"));

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
            Assert.Equal(1, properties.ChildElements.Count);

            Strike strike = properties.ChildElements[0] as Strike;
            Assert.NotNull(strike);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void StrikeS()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<s>test</s>"));

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
            Assert.Equal(1, properties.ChildElements.Count);

            Strike strike = properties.ChildElements[0] as Strike;
            Assert.NotNull(strike);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void UnderlineU()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<u>test</u>"));

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
            Assert.Equal(1, properties.ChildElements.Count);

            Underline underline = properties.ChildElements[0] as Underline;
            Assert.NotNull(underline);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void UnderlineIns()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ins>test</ins>"));

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
            Assert.Equal(1, properties.ChildElements.Count);

            Underline underline = properties.ChildElements[0] as Underline;
            Assert.NotNull(underline);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void UnderlineStrike()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<s><u>test</u></s>"));

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
            Assert.Equal(2, properties.ChildElements.Count);

            Strike underline = properties.ChildElements[0] as Strike;
            Assert.NotNull(underline);

            Underline strike = properties.ChildElements[1] as Underline;
            Assert.NotNull(strike);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TwoSpanWithSpace()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<span>one</span> <span>two</span>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(3, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("one", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(" ", text.InnerText);

            run = para.ChildElements[2] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TwoSpanWithMultiSpace()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<span>one</span>                <span>two</span>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(3, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("one", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(" ", text.InnerText);

            run = para.ChildElements[2] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TwoSpanWithSpaceInDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div><span>one</span> <span>two</span></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(3, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("one", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(" ", text.InnerText);

            run = para.ChildElements[2] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SpanWithTrailSpaceAndSpan()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<span>one </span> <span>two</span>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("one ", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SpanWithLeadingSpaceAndSpan()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<span>one</span> <span> two</span>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("one", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(" two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }
    }
}
