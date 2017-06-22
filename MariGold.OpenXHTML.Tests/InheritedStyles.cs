namespace MariGold.OpenXHTML.Tests
{
    using NUnit.Framework;
    using OpenXHTML;
    using System.IO;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Validation;
    using System.Linq;

    [TestFixture]
    public class InheritedStyles
    {
        [Test]
        public void MinimumStyle()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);
                doc.Process(new HtmlParser("<div style='font:10px verdana'><div style='font-family:arial'>test</div></div>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                OpenXmlElement para = doc.Document.Body.ChildElements[0];

                Assert.IsTrue(para is Paragraph);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                Assert.IsNotNull(run.RunProperties);
                Assert.AreEqual(2, run.RunProperties.ChildElements.Count);
                RunFonts fonts = run.RunProperties.ChildElements[0] as RunFonts;
                Assert.IsNotNull(fonts);
                Assert.AreEqual("arial", fonts.Ascii.Value);

                FontSize fontSize = run.RunProperties.ChildElements[1] as FontSize;
                Assert.IsNotNull(fontSize);
                Assert.AreEqual("20", fontSize.Val.Value);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void MaximumStyle()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);
                doc.Process(new HtmlParser("<div style='font:italic normal bold 10px/5px verdana'><div style='font-size:20px'>test</div></div>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                OpenXmlElement para = doc.Document.Body.ChildElements[0];

                Assert.IsTrue(para is Paragraph);
                Assert.AreEqual(2, para.ChildElements.Count);

                ParagraphProperties paraProperties = para.ChildElements[0] as ParagraphProperties;
                Assert.IsNotNull(paraProperties);
                SpacingBetweenLines space = paraProperties.ChildElements[0] as SpacingBetweenLines;
                Assert.IsNotNull(space);
                Assert.AreEqual("100", space.Line.Value);

                Run run = para.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                Assert.IsNotNull(run.RunProperties);
                Assert.AreEqual(4, run.RunProperties.ChildElements.Count);

                RunFonts fonts = run.RunProperties.ChildElements[0] as RunFonts;
                Assert.IsNotNull(fonts);
                Assert.AreEqual("verdana", fonts.Ascii.Value);

                Bold bold = run.RunProperties.ChildElements[1] as Bold;
                Assert.IsNotNull(bold);

                Italic italic = run.RunProperties.ChildElements[2] as Italic;
                Assert.IsNotNull(italic);

                FontSize fontSize = run.RunProperties.ChildElements[3] as FontSize;
                Assert.IsNotNull(fontSize);
                Assert.AreEqual("40", fontSize.Val.Value);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void H1AndSpan()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<h1><span>test</span></h1>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);
                var text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                Assert.IsNotNull(run.RunProperties);
                Bold bold = run.RunProperties.ChildElements[0] as Bold;
                Assert.IsNotNull(bold);
                FontSize fontSize = run.RunProperties.ChildElements[1] as FontSize;
                Assert.IsNotNull(fontSize);
                Assert.AreEqual("48", fontSize.Val.Value);


                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void H1AndSpanWithStyleSpan()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<h1><span style='font-size:48px'><span>test</span></span></h1>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);
                var text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                Assert.IsNotNull(run.RunProperties);
                Bold bold = run.RunProperties.ChildElements[0] as Bold;
                Assert.IsNotNull(bold);
                FontSize fontSize = run.RunProperties.ChildElements[1] as FontSize;
                Assert.IsNotNull(fontSize);
                Assert.AreEqual("96", fontSize.Val.Value);


                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }
    }
}
