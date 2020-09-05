namespace MariGold.OpenXHTML.Tests
{
    using System.IO;
    using System.Linq;
    using NUnit.Framework;
    using OpenXHTML;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Validation;

    [TestFixture]
    public class TestOL
    {
        [Test]
        public void TestSingle()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li></ol>"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("One", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void TestDouble()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li><li>Two</li></ol>"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void TestDoubleInBetweenDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li><div>One</div></li><li><div>Two</div></li></ol>"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void TwoElements()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li></ol><ol><li>Two</li></ol>"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void TwoElementsWithOneDouble()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li></ol><ol><li>Two</li><li>Three</li></ol>"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(3, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Two", text.InnerText);

            para = doc.Document.Body.ChildElements[2] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Three", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void ThreeElementsWithOneDouble()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li></ol><ol><li>Two</li><li>Three</li></ol><ol><li>Four</li></ol>"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(4, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Two", text.InnerText);

            para = doc.Document.Body.ChildElements[2] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Three", text.InnerText);

            para = doc.Document.Body.ChildElements[3] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(2, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Four", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void TwoElementsOneInsideAnother()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li><li><ol><li>Two</li></ol></li></ol>"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void TwoElementsOneInsideAnotherPInBetween()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li><li><p><ol><li>Two</li></ol></p></li></ol>"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.AreEqual(0, errors.Count());
        }

        [Test]
        public void ThreeElementsOneInsideAnotherPInBetween()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<ol><li>One</li><li><p><ol><li>Two<ol><li>Three</li></ol></li></ol></p></li></ol>"));

            Assert.IsNotNull(doc.Document.Body);
            Assert.AreEqual(3, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            ParagraphStyleId paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            NumberingProperties numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            NumberingLevelReference numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(0, numberingLevelReference.Val.Value);

            NumberingId numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("One", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(1, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Two", text.InnerText);

            para = doc.Document.Body.ChildElements[2] as Paragraph;

            Assert.IsNotNull(para);
            Assert.AreEqual(2, para.ChildElements.Count);

            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.IsNotNull(properties);
            Assert.AreEqual(2, properties.ChildElements.Count);
            paragraphStyleId = properties.ChildElements[0] as ParagraphStyleId;
            Assert.IsNotNull(paragraphStyleId);
            Assert.AreEqual("ListParagraph", paragraphStyleId.Val.Value);

            numberingProperties = properties.ChildElements[1] as NumberingProperties;
            Assert.IsNotNull(numberingProperties);
            Assert.AreEqual(2, numberingProperties.ChildElements.Count);

            numberingLevelReference = numberingProperties.ChildElements[0] as NumberingLevelReference;
            Assert.IsNotNull(numberingLevelReference);
            Assert.AreEqual(2, numberingLevelReference.Val.Value);

            numberingId = numberingProperties.ChildElements[1] as NumberingId;
            Assert.IsNotNull(numberingId);
            Assert.AreEqual(1, numberingId.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.IsNotNull(run);
            Assert.AreEqual(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.IsNotNull(text);
            Assert.AreEqual("Three", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.AreEqual(0, errors.Count());
        }
    }
}
