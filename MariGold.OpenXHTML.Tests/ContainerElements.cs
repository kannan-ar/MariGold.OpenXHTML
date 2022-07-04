namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class ContainerElements
    {
        [Fact]
        public void SimpleUl()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ul><li>test</li></ul>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(2, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void UlWithH1()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ul><li><h1>test</h1></li></ul>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(2, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties runProperties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(runProperties);
            Bold bold = runProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);
            FontSize fontSize = runProperties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("48", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SimpleOL()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>test</li></ol>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DoubleOL()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>one</li></ol><ol><li>two</li></ol>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("one", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(1, numberingId.Val.Value);

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
        public void OlInsideUl()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ul><li>One<ol><li>Two</li></ol></li></ul>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(2, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(0, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("Two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void OlInsideUlInsideUl()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ul><li>One<ol><li>Two<ul><li>Three</li></ul></li></ol></li></ul>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(3, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(2, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(0, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("Two", text.InnerText);

            para = doc.Document.Body.ChildElements[2] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(2, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(3, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("Three", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void OlInsideUlInsideUlBetweenPAndDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ul><li>One<div><ol><li>Two<p><ul><li>Three</li></ul></p></li></ol></div></li></ul>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(3, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(2, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(0, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("Two", text.InnerText);

            para = doc.Document.Body.ChildElements[2] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.NotNull(paragraphStyleId);
            Assert.Equal("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.NotNull(numberingProperties);
            Assert.Equal(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.NotNull(numberingLevelReference);
            Assert.Equal(2, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(3, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("Three", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }
    }
}
