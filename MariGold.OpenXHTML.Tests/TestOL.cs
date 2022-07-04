namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class TestOL
    {
        [Fact]
        public void TestSingle()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li></ol>"));

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
            Assert.Equal("One", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TestDouble()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li><li>Two</li></ol>"));

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
        public void TestDoubleInBetweenDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li><div>One</div></li><li><div>Two</div></li></ol>"));

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
        public void TwoElements()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li></ol><ol><li>Two</li></ol>"));

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
            Assert.Equal(1, numberingLevelReference.Val.Value);

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
        public void TwoElementsWithOneDouble()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li></ol><ol><li>Two</li><li>Three</li></ol>"));

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
            Assert.Equal(1, numberingId.Val.Value);

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
            Assert.Equal(1, numberingLevelReference.Val.Value);

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
            Assert.Equal(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(1, numberingId.Val.Value);

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
        public void ThreeElementsWithOneDouble()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li></ol><ol><li>Two</li><li>Three</li></ol><ol><li>Four</li></ol>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(4, doc.Document.Body.ChildElements.Count);

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
            Assert.Equal(1, numberingLevelReference.Val.Value);

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
            Assert.Equal(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.NotNull(numberingId);
            Assert.Equal(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("Three", text.InnerText);

            para = doc.Document.Body.ChildElements[3] as Paragraph;

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
            Assert.Equal(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("Four", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TwoElementsOneInsideAnother()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li><li><ol><li>Two</li></ol></li></ol>"));

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
            Assert.Equal(1, numberingLevelReference.Val.Value);

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
        public void TwoElementsOneInsideAnotherPInBetween()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li><li><p><ol><li>Two</li></ol></p></li></ol>"));

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
            Assert.Equal(1, numberingLevelReference.Val.Value);

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
        public void ThreeElementsOneInsideAnotherPInBetween()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li><li><p><ol><li>Two<ol><li>Three</li></ol></li></ol></p></li></ol>"));

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
            Assert.Equal(1, numberingId.Val.Value);

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
            Assert.Equal(1, numberingLevelReference.Val.Value);

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
            Assert.Equal(1, numberingId.Val.Value);

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
