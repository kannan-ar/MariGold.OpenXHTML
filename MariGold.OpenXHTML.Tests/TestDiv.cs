namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class TestDiv
    {
        [Fact]
        public void SingleDivPercentageFontSize()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='font-size:100%'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            FontSize fontSize = run.RunProperties.ChildElements[0] as FontSize;
            Assert.Equal("32", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SingleDivOneEmFontSize()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='font-size:1em'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            FontSize fontSize = run.RunProperties.ChildElements[0] as FontSize;
            Assert.Equal("24", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SingleDivXXLargeFontSize()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='font-size:xx-large'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            FontSize fontSize = run.RunProperties.ChildElements[0] as FontSize;
            Assert.Equal("48", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void MarginDivAndWidthoutMarginDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='margin:5px'>1</div><div>2</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(2, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties paragraphProperties = paragraph.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(paragraphProperties);
            Assert.Equal(2, paragraphProperties.ChildElements.Count);
            SpacingBetweenLines spacing = paragraphProperties.ChildElements[0] as SpacingBetweenLines;
            Assert.NotNull(spacing);
            Assert.Equal("100", spacing.Before.Value);
            Assert.Equal("100", spacing.After.Value);
            Indentation ind = paragraphProperties.ChildElements[1] as Indentation;
            Assert.NotNull(ind);
            Assert.Equal("100", ind.Left.Value);
            Assert.Equal("100", ind.Right.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("1", text.InnerText);

            paragraph = doc.Document.Body.ChildElements[1] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(1, paragraph.ChildElements.Count);

            run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("2", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void ParagraphLineHeight()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='line-height:50px'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties properties = paragraph.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            SpacingBetweenLines space = properties.ChildElements[0] as SpacingBetweenLines;
            Assert.NotNull(space);
            Assert.Equal("1000", space.Line.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void DivInATag()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com'><div>test</div></a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink hyperLink = para.ChildElements[0] as Hyperlink;
            Assert.NotNull(hyperLink);
            Assert.Equal(1, hyperLink.ChildElements.Count);

            Run run = hyperLink.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);
        }

        [Fact]
        public void ParagraphNormalLineHeight()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='line-height:normal'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(1, paragraph.ChildElements.Count);

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void InvalidMargin()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='margin:hdh'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(1, paragraph.ChildElements.Count);

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ParagraphNumberLineHeight()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='line-height:5'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties properties = paragraph.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            SpacingBetweenLines space = properties.ChildElements[0] as SpacingBetweenLines;
            Assert.NotNull(space);
            Assert.Equal("1280", space.Line.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void PercentageEm()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='margin-bottom:.35em'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;

            ParagraphProperties paragraphProperties = paragraph.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(paragraphProperties);
            Assert.Equal(1, paragraphProperties.ChildElements.Count);
            SpacingBetweenLines spacing = paragraphProperties.ChildElements[0] as SpacingBetweenLines;
            Assert.NotNull(spacing);
            Assert.Equal("84", spacing.After.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ChildBackgroundTransparent()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='background:#000'><div style='background-color:transparent'>test</div></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);

            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties properties = paragraph.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.NotNull(properties.Shading);
            Assert.Equal("000000", properties.Shading.Fill.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ParagraphDecimalLineHeight()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='line-height:1.5'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties properties = paragraph.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            SpacingBetweenLines space = properties.ChildElements[0] as SpacingBetweenLines;
            Assert.NotNull(space);
            Assert.Equal("160", space.Line.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void PageBreak()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='page-break-before:always'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties properties = paragraph.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.NotNull(properties.PageBreakBefore);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void PageBreakWithTwoDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div>1</div><div style='page-break-before:always'>2</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(2, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(1, paragraph.ChildElements.Count);
            Assert.Null(paragraph.ParagraphProperties);

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.Equal("1", text.InnerText);

            paragraph = doc.Document.Body.ChildElements[1] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties properties = paragraph.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.NotNull(properties.PageBreakBefore);

            run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            text = run.ChildElements[0] as Word.Text;
            Assert.Equal("2", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }
    }
}
