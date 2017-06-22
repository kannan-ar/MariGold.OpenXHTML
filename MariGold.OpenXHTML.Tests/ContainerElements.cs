namespace MariGold.OpenXHTML.Tests
{
    using System;
    using NUnit.Framework;
    using OpenXHTML;
    using System.IO;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Validation;
    using System.Linq;

    [TestFixture]
    public class ContainerElements
    {
        [Test]
        public void SimpleUl()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<ul><li>test</li></ul>"));

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
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void UlWithH1()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<ul><li><h1>test</h1></li></ul>"));

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
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties runProperties = run.ChildElements[0] as RunProperties;
                Assert.IsNotNull(runProperties);
                Bold bold = runProperties.ChildElements[0] as Bold;
                Assert.IsNotNull(bold);
                FontSize fontSize = runProperties.ChildElements[1] as FontSize;
                Assert.IsNotNull(fontSize);
                Assert.AreEqual("48", fontSize.Val.Value);

                Word.Text text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void SimpleOL()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<ol><li>test</li></ol>"));

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
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void DoubleOL()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<ol><li>one</li></ol><ol><li>two</li></ol>"));

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
                Assert.AreEqual("one", text.InnerText);

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
                Assert.AreEqual(2, numberingId.Val.Value);

                run = para.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("two", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }
    }
}
