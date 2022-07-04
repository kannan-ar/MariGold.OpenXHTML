namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class HtmlDefaultStyles
    {
        [Fact]
        public void ThWithSpanAndDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><th><span>one</span><div>two</div></th></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableGrid tableGrid = table.ChildElements[1] as TableGrid;
            Assert.NotNull(tableGrid);
            Assert.Equal(1, tableGrid.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(2, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            Bold bold = run.RunProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("one", text.InnerText);

            para = cell.ChildElements[1] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            bold = run.RunProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivInsideItalic()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<i><div>test</div></i>"));

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

            Italic italic = properties.ChildElements[0] as Italic;
            Assert.NotNull(italic);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TwoLevelDivInsideItalic()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<i><div><div>test</div></div></i>"));

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

            Italic italic = properties.ChildElements[0] as Italic;
            Assert.NotNull(italic);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void DivInsideB()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<b><div>test</div></b>"));

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

            Bold bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TwoLevelDivInsideB()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<b><div><div>test</div></div></b>"));

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

            Bold bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TwoLevelDivInsideH1()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<h1><div>test</div></h1>"));

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
            Assert.Equal(2, properties.ChildElements.Count);

            Bold bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            FontSize fontSize = properties.ChildElements[1] as FontSize;
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
        public void TwoLevelDivInsideUnderline()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<u><div><div>test</div></div></u>"));

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

            Underline underline = properties.ChildElements[0] as Underline;
            Assert.NotNull(underline);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void BInsideNoFontWeightDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style=\"font-weight:normal\"><b>test</b></div>"));

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

            Bold bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void IInsideNoTextDecorationDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style=\"text-decoration:none\"><u>test</u></div>"));

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

            Underline underline = properties.ChildElements[0] as Underline;
            Assert.NotNull(underline);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void AddressInsideFontStyleNormalDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style=\"font-style:normal\"><address>test</address></div>"));

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

            Italic italic = properties.ChildElements[0] as Italic;
            Assert.NotNull(italic);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void CenterInsideTextAlignLeftDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style=\"text-align:left\"><center>test</center></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            ParagraphProperties paraProperties = para.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(paraProperties);
            Assert.Equal(1, paraProperties.ChildElements.Count);
            Justification justification = paraProperties.ChildElements[0] as Justification;
            Assert.NotNull(justification);
            Assert.Equal(JustificationValues.Center, justification.Val.Value);

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
        public void H1DefaultStyleIn10PxDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style=\"font-size:10px\"><h1>test</h1></div>"));

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
            Assert.Equal(2, properties.ChildElements.Count);

            Bold bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            FontSize fontSize = properties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("40", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void H1DefaultStyleIn1EmDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style=\"font-size:1em\"><div><h1>test</h1></div></div>"));

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
            Assert.Equal(2, properties.ChildElements.Count);

            Bold bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            FontSize fontSize = properties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("64", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void H1DefaultStyleIn50PercentageDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style=\"font-size:50%\"><div><h1>test</h1></div></div>"));

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
            Assert.Equal(2, properties.ChildElements.Count);

            Bold bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            FontSize fontSize = properties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("32", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void H3DefaultStyleIn10PxDiv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style=\"font-size:10px\"><div><h3>test</h3></div></div>"));

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
            Assert.Equal(2, properties.ChildElements.Count);

            Bold bold = properties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            FontSize fontSize = properties.ChildElements[1] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("23", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }
    }
}
