namespace MariGold.OpenXHTML.Tests
{
    using NUnit.Framework;
    using OpenXHTML;
    using System.IO;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Validation;
    using System.Linq;

    [TestFixture]
    public class InlineElements
    {
        [Test]
        public void FontTagDefault()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<font>test</font>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);
                var text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void SpanDefault()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<span>test</span>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);
                var text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void FontTagWithAttributes()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<font face=\"Arial\" size=\"3\" color=\"red\">test</font>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;
                Assert.IsNotNull(properties);
                Assert.AreEqual(3, properties.ChildElements.Count);

                RunFonts fonts = properties.ChildElements[0] as RunFonts;
                Assert.IsNotNull(fonts);
                Assert.AreEqual("Arial", fonts.Ascii.Value);

                Word.Color color = properties.ChildElements[1] as Word.Color;
                Assert.AreEqual("FF0000", color.Val.Value);

                FontSize fontSize = properties.ChildElements[2] as FontSize;
                Assert.IsNotNull(fontSize);
                Assert.AreEqual("32", fontSize.Val.Value);

                var text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void TestElement()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<test>one</test>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                var text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("one", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void QuoteElement()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<q>one</q>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(3, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                var text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("\"", text.InnerText);

                run = para.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("one", text.InnerText);

                run = para.ChildElements[2] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("\"", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void SpanQuoteElement()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<span>test</span><q>one</q>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(4, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                var text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                run = para.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("\"", text.InnerText);

                run = para.ChildElements[2] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("one", text.InnerText);

                run = para.ChildElements[3] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("\"", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void Superscript()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<sup>test</sup>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;
                Assert.IsNotNull(properties);
                Assert.AreEqual(1, properties.ChildElements.Count);

                VerticalTextAlignment verticalTextAlignment = properties.ChildElements[0] as VerticalTextAlignment;
                Assert.IsNotNull(verticalTextAlignment);
                Assert.AreEqual(VerticalPositionValues.Superscript, verticalTextAlignment.Val.Value);

                Word.Text text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void Subscript()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<sub>test</sub>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
                
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;
                Assert.IsNotNull(properties);
                Assert.AreEqual(1, properties.ChildElements.Count);

                VerticalTextAlignment verticalTextAlignment = properties.ChildElements[0] as VerticalTextAlignment;
                Assert.IsNotNull(verticalTextAlignment);
                Assert.AreEqual(VerticalPositionValues.Subscript, verticalTextAlignment.Val.Value);

                Word.Text text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void StrikeDel()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<del>test</del>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;
                Assert.IsNotNull(properties);
                Assert.AreEqual(1, properties.ChildElements.Count);

                Strike strike = properties.ChildElements[0] as Strike;
                Assert.IsNotNull(strike);

                Word.Text text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void StrikeS()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<s>test</s>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;
                Assert.IsNotNull(properties);
                Assert.AreEqual(1, properties.ChildElements.Count);

                Strike strike = properties.ChildElements[0] as Strike;
                Assert.IsNotNull(strike);

                Word.Text text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void UnderlineU()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<u>test</u>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;
                Assert.IsNotNull(properties);
                Assert.AreEqual(1, properties.ChildElements.Count);

                Underline underline = properties.ChildElements[0] as Underline;
                Assert.IsNotNull(underline);

                Word.Text text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void UnderlineIns()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<ins>test</ins>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;
                Assert.IsNotNull(properties);
                Assert.AreEqual(1, properties.ChildElements.Count);

                Underline underline = properties.ChildElements[0] as Underline;
                Assert.IsNotNull(underline);

                Word.Text text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void UnderlineStrike()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<s><u>test</u></s>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;
                Assert.IsNotNull(properties);
                Assert.AreEqual(2, properties.ChildElements.Count);

                Strike underline = properties.ChildElements[0] as Strike;
                Assert.IsNotNull(underline);

                Underline strike = properties.ChildElements[1] as Underline;
                Assert.IsNotNull(strike);

                Word.Text text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void TwoSpanWithSpace()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<span>one</span> <span>two</span>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(3, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("one", text.InnerText);

                run = para.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual(" ", text.InnerText);

                run = para.ChildElements[2] as Run;
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

        [Test]
        public void TwoSpanWithMultiSpace()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<span>one</span>                <span>two</span>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(3, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("one", text.InnerText);

                run = para.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual(" ", text.InnerText);

                run = para.ChildElements[2] as Run;
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

        [Test]
        public void TwoSpanWithSpaceInDiv()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<div><span>one</span> <span>two</span></div>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(3, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("one", text.InnerText);

                run = para.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual(" ", text.InnerText);

                run = para.ChildElements[2] as Run;
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

        [Test]
        public void SpanWithTrailSpaceAndSpan()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<span>one </span> <span>two</span>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(2, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("one ", text.InnerText);

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

        [Test]
        public void SpanWithLeadingSpaceAndSpan()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<span>one</span> <span> two</span>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(2, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("one", text.InnerText);

                run = para.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual(" two", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }
    }
}
