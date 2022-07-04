namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class BasicDocument
    {
        [Fact]
        public void EmptyString()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser(" "));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(0, doc.Document.Body.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);

        }

        [Fact]
        public void EmptyBody()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<body></body>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(0, doc.Document.Body.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SimpleText()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("test"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            OpenXmlElement para = doc.Document.Body.ChildElements[0];

            Assert.True(para is Paragraph);
            Assert.Equal(1, para.ChildElements.Count);

            OpenXmlElement run = para.ChildElements[0];
            Assert.True(run is Run);
            Assert.Equal(1, run.ChildElements.Count);

            OpenXmlElement text = run.ChildElements[0] as DocumentFormat.OpenXml.Wordprocessing.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SimpleTable()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><td>1</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("1", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TwoSpanOnBody()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<span>1</span><span>2</span>"));

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
            Assert.Equal("1", text.InnerText);

            run = para.ChildElements[1] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("2", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivAndSpanOnly()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div>1</div><span>2</span>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("1", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("2", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SpanAndDivOnly()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<span>1</span><div>2</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("1", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("2", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivAndDivOnly()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div>1</div><div>2</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("1", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("2", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivTwoSpanOnBody()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div>1</div><span>2</span><span>3</span>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("1", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;
            text = run.ChildElements[0] as Word.Text;
            Assert.Equal("2", text.InnerText);

            run = para.ChildElements[1] as Run;
            text = run.ChildElements[0] as Word.Text;
            Assert.Equal("3", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void OneAOnBody()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com'>click here</a>"));

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

            RunStyle runStyle = properties.ChildElements[0] as RunStyle;

            Assert.NotNull(runStyle);
            Assert.Equal("Hyperlink", runStyle.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("click here", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void AOnDivBody()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div><a href='http://google.com'>click here</a></div>"));

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

            RunStyle runStyle = properties.ChildElements[0] as RunStyle;

            Assert.NotNull(runStyle);
            Assert.Equal("Hyperlink", runStyle.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("click here", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TextAndSpanOnDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div>pp<span>test1</span></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);
            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            var text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("pp", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test1", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivInsideAnotherDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div><div>test</div></div>"));

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
        public void TextWithBreak()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("test<br />text"));

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
            Assert.Equal("test", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            var br = run.ChildElements[0] as Break;
            Assert.NotNull(br);

            run = para.ChildElements[2] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("text", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivTextWithBreak()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div>test<br />text</div>"));

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
            Assert.Equal("test", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            var br = run.ChildElements[0] as Break;
            Assert.NotNull(br);

            run = para.ChildElements[2] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("text", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivSpanTextWithBreak()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div><span>test</span><br /><span>text</span></div>"));

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
            Assert.Equal("test", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            var br = run.ChildElements[0] as Break;
            Assert.NotNull(br);

            run = para.ChildElements[2] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("text", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivSpanStyleTextWithBreak()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div><span style='color:#ff0000'>test</span><br /><span>text</span></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(3, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            var text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            Assert.NotNull(run.RunProperties);
            Assert.NotNull(run.RunProperties.Color);
            Assert.Equal("ff0000", run.RunProperties.Color.Val.Value);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            var br = run.ChildElements[0] as Break;
            Assert.NotNull(br);

            run = para.ChildElements[2] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("text", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void InnerDivAndText()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div><div>one</div>two</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(2, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            var text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("one", text.InnerText);

            para = doc.Document.Body.ChildElements[1] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;
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
        public void H1Only()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<h1>test</h1>"));

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
        public void SimpleAddress()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<address>first line<br />second line</address>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            OpenXmlElement para = doc.Document.Body.ChildElements[0];

            Assert.True(para is Paragraph);
            Assert.Equal(3, para.ChildElements.Count);

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
            Assert.Equal("first line", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Break br = run.ChildElements[1] as Break;
            Assert.NotNull(br);

            run = para.ChildElements[2] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            Assert.NotNull(run.RunProperties);
            Assert.Equal(1, run.RunProperties.ChildElements.Count);
            italic = run.RunProperties.ChildElements[0] as Italic;
            Assert.NotNull(italic);

            text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("second line", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SimpleDL()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<dl><dt>Numbers</dt><dd>1</dd><dt>Text</dt><dd>One</dd></dl>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(6, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.Equal(1, para.ChildElements.Count);
            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);
            SpacingBetweenLines spacing = properties.ChildElements[0] as SpacingBetweenLines;
            Assert.NotNull(spacing);
            Assert.Equal("240", spacing.Before.Value);

            para = doc.Document.Body.ChildElements[1] as Paragraph;
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("Numbers", text.InnerText);

            para = doc.Document.Body.ChildElements[2] as Paragraph;
            Assert.Equal(2, para.ChildElements.Count);
            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);
            Indentation ind = properties.ChildElements[0] as Indentation;
            Assert.NotNull(ind);
            Assert.Equal("800", ind.Left.Value);
            Assert.Null(ind.Right);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("1", text.InnerText);

            para = doc.Document.Body.ChildElements[3] as Paragraph;
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("Text", text.InnerText);

            para = doc.Document.Body.ChildElements[4] as Paragraph;
            Assert.Equal(2, para.ChildElements.Count);
            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);
            ind = properties.ChildElements[0] as Indentation;
            Assert.NotNull(ind);
            Assert.Equal("800", ind.Left.Value);
            Assert.Null(ind.Right);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("One", text.InnerText);

            para = doc.Document.Body.ChildElements[5] as Paragraph;
            Assert.Equal(1, para.ChildElements.Count);
            properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            spacing = properties.ChildElements[0] as SpacingBetweenLines;
            Assert.NotNull(spacing);
            Assert.Null(spacing.Before);
            Assert.Equal("240", spacing.After.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void OnlyHr()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<hr />"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.Equal(2, para.ChildElements.Count);
            ParagraphProperties properties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            ParagraphBorders borders = properties.ChildElements[0] as ParagraphBorders;
            Assert.NotNull(borders);
            Assert.Equal(1, borders.ChildElements.Count);

            TopBorder topBorder = borders.ChildElements[0] as TopBorder;
            Assert.NotNull(topBorder);
            TestUtility.TestBorder<TopBorder>(topBorder, BorderValues.Single, "auto", 4U);

            Run run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(string.Empty, text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ATag()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com'>test</a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink hyperLink = para.ChildElements[0] as Hyperlink;
            Assert.NotNull(hyperLink);

            Run run = hyperLink.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            RunStyle runStyle = properties.ChildElements[0] as RunStyle;
            Assert.NotNull(runStyle);
            Assert.Equal("Hyperlink", runStyle.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ATagWithBold()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com'><strong>bold</strong>test</a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink hyperLink = para.ChildElements[0] as Hyperlink;
            Assert.NotNull(hyperLink);

            Run run = hyperLink.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            Bold bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("bold", text.InnerText);

            run = hyperLink.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            RunStyle runStyle = properties.ChildElements[0] as RunStyle;
            Assert.NotNull(runStyle);
            Assert.Equal("Hyperlink", runStyle.Val.Value);

            text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void SpanInsideATag()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com'><span>click</span> here</a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;

            Assert.NotNull(link);
            Assert.Equal(2, link.ChildElements.Count);

            Run run = link.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("click", text.InnerText);

            run = link.ChildElements[1] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;

            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            RunStyle runStyle = properties.ChildElements[0] as RunStyle;

            Assert.NotNull(runStyle);
            Assert.Equal("Hyperlink", runStyle.Val.Value);

            text = run.ChildElements[1] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal(" here", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void InsideATagTextSpan()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com'>here<span>click</span></a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;

            Assert.NotNull(link);
            Assert.Equal(2, link.ChildElements.Count);

            Run run = link.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;

            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            RunStyle runStyle = properties.ChildElements[0] as RunStyle;

            Assert.NotNull(runStyle);
            Assert.Equal("Hyperlink", runStyle.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("here", text.InnerText);

            run = link.ChildElements[1] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("click", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void InsideATagTextBr()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com'>here<br /></a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;

            Assert.NotNull(link);
            Assert.Equal(2, link.ChildElements.Count);

            Run run = link.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;

            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            RunStyle runStyle = properties.ChildElements[0] as RunStyle;

            Assert.NotNull(runStyle);
            Assert.Equal("Hyperlink", runStyle.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("here", text.InnerText);

            run = link.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Break br = run.ChildElements[0] as Break;
            Assert.NotNull(br);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void InsideATagCenterText()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com'><center>click </center>here</a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;

            Assert.NotNull(link);
            Assert.Equal(2, link.ChildElements.Count);

            Run run = link.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("click ", text.InnerText);

            run = link.ChildElements[1] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;

            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            RunStyle runStyle = properties.ChildElements[0] as RunStyle;

            Assert.NotNull(runStyle);
            Assert.Equal("Hyperlink", runStyle.Val.Value);

            text = run.ChildElements[1] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("here", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void InsideATagItalicText()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com'><i>click </i>here</a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;

            Assert.NotNull(link);
            Assert.Equal(2, link.ChildElements.Count);

            Run run = link.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties runProperties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(runProperties);
            Assert.Equal(1, runProperties.ChildElements.Count);
            Italic italic = runProperties.ChildElements[0] as Italic;
            Assert.NotNull(italic);

            Word.Text text = run.ChildElements[1] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("click ", text.InnerText);

            run = link.ChildElements[1] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;

            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            RunStyle runStyle = properties.ChildElements[0] as RunStyle;

            Assert.NotNull(runStyle);
            Assert.Equal("Hyperlink", runStyle.Val.Value);

            text = run.ChildElements[1] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal("here", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TwoATagWithSpan()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<a href='http://google.com'><span>one</span></a><a href='#'><span>two</span></a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            Hyperlink link = para.ChildElements[0] as Hyperlink;

            Assert.NotNull(link);
            Assert.Equal(1, link.ChildElements.Count);

            Run run = link.ChildElements[0] as Run;

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
            Assert.Equal("two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ProtocolFreeUrl()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            doc.UriSchema = Uri.UriSchemeHttp;
            doc.Process(new HtmlParser("<a href='//google.com'>test</a>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Hyperlink hyperLink = para.ChildElements[0] as Hyperlink;
            Assert.NotNull(hyperLink);

            Run run = hyperLink.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            Assert.Equal(1, properties.ChildElements.Count);

            RunStyle runStyle = properties.ChildElements[0] as RunStyle;
            Assert.NotNull(runStyle);
            Assert.Equal("Hyperlink", runStyle.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void DisplayNone()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style='display:none'>test</div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(0, doc.Document.Body.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void H1ToH6DefaultStyles()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<h1>test</h1><h2>test</h2><h3>test</h3><h4>test</h4><h5>test</h5><h6>test</h6>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(6, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);
            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            RunProperties properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            Bold bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);
            FontSize fontSize = properties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("48", fontSize.Val.Value);

            para = doc.Document.Body.ChildElements[1] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);
            run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);
            fontSize = properties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("36", fontSize.Val.Value);

            para = doc.Document.Body.ChildElements[2] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);
            run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);
            fontSize = properties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("28", fontSize.Val.Value);

            para = doc.Document.Body.ChildElements[3] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);
            run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);
            fontSize = properties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("24", fontSize.Val.Value);

            para = doc.Document.Body.ChildElements[4] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);
            run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);
            fontSize = properties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("20", fontSize.Val.Value);

            para = doc.Document.Body.ChildElements[5] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);
            run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);
            bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);
            fontSize = properties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("16", fontSize.Val.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void SectionElement()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<section>test</section>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);
            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }
    }
}
