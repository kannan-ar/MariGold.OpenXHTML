namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class InheritedStyles
    {
        [Fact]
        public void MinimumStyle()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            doc.Process(new HtmlParser("<div style='font:10px verdana'><div style='font-family:arial'>test</div></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            OpenXmlElement para = doc.Document.Body.ChildElements[0];

            Assert.True(para is Paragraph);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Assert.NotNull(run.RunProperties);
            Assert.Equal(2, run.RunProperties.ChildElements.Count);
            RunFonts fonts = run.RunProperties.ChildElements[0] as RunFonts;
            Assert.NotNull(fonts);
            Assert.Equal("arial", fonts.Ascii.Value);

            FontSize fontSize = run.RunProperties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("20", fontSize.Val.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void MaximumStyle()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            doc.Process(new HtmlParser("<div style='font:italic normal bold 10px/5px verdana'><div style='font-size:20px'>test</div></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            OpenXmlElement para = doc.Document.Body.ChildElements[0];

            Assert.True(para is Paragraph);
            Assert.Equal(2, para.ChildElements.Count);

            ParagraphProperties paraProperties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(paraProperties);
            SpacingBetweenLines space = paraProperties.ChildElements[0] as SpacingBetweenLines;
            Assert.NotNull(space);
            Assert.Equal("100", space.Line.Value);

            Run run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Assert.NotNull(run.RunProperties);
            Assert.Equal(4, run.RunProperties.ChildElements.Count);

            RunFonts fonts = run.RunProperties.ChildElements[0] as RunFonts;
            Assert.NotNull(fonts);
            Assert.Equal("verdana", fonts.Ascii.Value);

            Bold bold = run.RunProperties.ChildElements[1] as Bold;
            Assert.NotNull(bold);

            Italic italic = run.RunProperties.ChildElements[2] as Italic;
            Assert.NotNull(italic);

            FontSize fontSize = run.RunProperties.ChildElements[3] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("40", fontSize.Val.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void H1AndSpan()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<h1><span>test</span></h1>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            var text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            Assert.NotNull(run.RunProperties);
            Bold bold = run.RunProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);
            FontSize fontSize = run.RunProperties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("48", fontSize.Val.Value);


            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void H1AndSpanWithStyleSpan()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<h1><span style='font-size:48px'><span>test</span></span></h1>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            var text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            Assert.NotNull(run.RunProperties);
            Bold bold = run.RunProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);
            FontSize fontSize = run.RunProperties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("96", fontSize.Val.Value);


            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }
    }
}
