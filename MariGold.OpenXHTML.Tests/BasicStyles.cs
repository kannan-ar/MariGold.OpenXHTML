namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class BasicStyles
    {
        [Fact]
        public void DivColorRed()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='color:#ff0000'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            OpenXmlElement para = doc.Document.Body.ChildElements[0];

            Assert.True(para is Paragraph);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Assert.NotNull(run.RunProperties);
            Assert.NotNull(run.RunProperties.Color);
            Assert.Equal("ff0000", run.RunProperties.Color.Val.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivRGBColorRed()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='color:rgb(255,0,0)'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            OpenXmlElement para = doc.Document.Body.ChildElements[0];

            Assert.True(para is Paragraph);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Assert.NotNull(run.RunProperties);
            Assert.NotNull(run.RunProperties.Color);
            Assert.Equal("FF0000", run.RunProperties.Color.Val.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void iTag()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<i>test</i>"));

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
            Italic italic = run.RunProperties.ChildElements[0] as Italic;
            Assert.NotNull(italic);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivUnderline()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='text-decoration:underline'>test</div>"));

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
            Underline underline = run.RunProperties.ChildElements[0] as Underline;
            Assert.NotNull(underline);
            Assert.Equal(UnderlineValues.Single, underline.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);

        }

        [Fact]
        public void DivTextDecorationLine()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='text-decoration-line:underline'>test</div>"));

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
            Underline underline = run.RunProperties.ChildElements[0] as Underline;
            Assert.NotNull(underline);
            Assert.Equal(UnderlineValues.Single, underline.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);

        }

        [Fact]
        public void BTag()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<b>test</b>"));

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
        public void StrongTagWithSpace()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<strong>Name &amp; SSN </string>"));

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
            Assert.Equal("Name & SSN ", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void BackgroundProperty()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='background:#000 no-repeat right top;'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.Equal(2, paragraph.ChildElements.Count);
            Assert.NotNull(paragraph.ParagraphProperties);
            Assert.NotNull(paragraph.ParagraphProperties.Shading);
            Assert.Equal("000000", paragraph.ParagraphProperties.Shading.Fill.Value);
            Assert.Equal(Word.ShadingPatternValues.Clear, paragraph.ParagraphProperties.Shading.Val.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void FontSizeOnInnerSpan()
        {
            string html = "<a href=\"http://google.com\" style='font-size:24px'><span>click here</span></a>";

            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            doc.Process(new HtmlParser(html));

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
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;

            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            FontSize fontSize = properties.ChildElements[0] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("48", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("click here", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void HeaderTagStyleOverride()
        {
            string html = "<h2 style='font-size:10px;font-weight:normal'>test</h2>";

            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            doc.Process(new HtmlParser(html));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);

            Assert.Equal(2, run.ChildElements.Count);

            RunProperties runProperties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(runProperties);
            Assert.Equal(1, runProperties.ChildElements.Count);

            FontSize fontSize = runProperties.ChildElements[0] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("20", fontSize.Val.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SpanSuperscriptStyle()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<span style=\"vertical-align:super\">test</span>"));

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
        public void SpanSubscriptStyle()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<span style=\"vertical-align:sub\">test</span>"));

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
        public void FiftyPercentageEMFontSize()
        {
            string html = "<div style=\"font-size:0.50em\">test</div>";

            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            doc.Process(new HtmlParser(html));

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

            FontSize fontSize = properties.ChildElements[0] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("12", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }
    }
}
