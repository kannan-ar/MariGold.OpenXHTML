namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class TestImage
    {
        [Fact]
        public void ImageWithoutWidthAndHeight()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div><img src=\"data:image/png,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" /></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);
            Assert.Equal(1, para.ChildElements.Count);

            DW.Inline inline = image.ChildElements[0] as DW.Inline;
            Assert.NotNull(inline);
            Assert.Equal(5, inline.ChildElements.Count);

            DW.Extent extent = inline.ChildElements[0] as DW.Extent;
            Assert.NotNull(extent);
            Assert.Equal(47625, extent.Cx);
            Assert.Equal(47625, extent.Cy);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ImageWithOwnWidth()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div><img src=\"data:image/png,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" style=\"width:10px\" /></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);
            Assert.Equal(1, para.ChildElements.Count);

            DW.Inline inline = image.ChildElements[0] as DW.Inline;
            Assert.NotNull(inline);
            Assert.Equal(5, inline.ChildElements.Count);

            DW.Extent extent = inline.ChildElements[0] as DW.Extent;
            Assert.NotNull(extent);
            Assert.Equal(95250, extent.Cx);
            Assert.Equal(95250, extent.Cy);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ImageWithInheritedWidth()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style=\"width:10px\"><img src=\"data:image/png,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" /></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);
            Assert.Equal(1, para.ChildElements.Count);

            DW.Inline inline = image.ChildElements[0] as DW.Inline;
            Assert.NotNull(inline);
            Assert.Equal(5, inline.ChildElements.Count);

            DW.Extent extent = inline.ChildElements[0] as DW.Extent;
            Assert.NotNull(extent);
            Assert.Equal(95250, extent.Cx);
            Assert.Equal(95250, extent.Cy);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ImageWithOwnHeight()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div><img src=\"data:image/png,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" style=\"height:10px\" /></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);
            Assert.Equal(1, para.ChildElements.Count);

            DW.Inline inline = image.ChildElements[0] as DW.Inline;
            Assert.NotNull(inline);
            Assert.Equal(5, inline.ChildElements.Count);

            DW.Extent extent = inline.ChildElements[0] as DW.Extent;
            Assert.NotNull(extent);
            Assert.Equal(47625, extent.Cx);
            Assert.Equal(95250, extent.Cy);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ImageWithInheritedHeight()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<div style=\"height:10px\"><img src=\"data:image/png,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" /></div>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);
            Assert.Equal(1, para.ChildElements.Count);

            DW.Inline inline = image.ChildElements[0] as DW.Inline;
            Assert.NotNull(inline);
            Assert.Equal(5, inline.ChildElements.Count);

            DW.Extent extent = inline.ChildElements[0] as DW.Extent;
            Assert.NotNull(extent);
            Assert.Equal(47625, extent.Cx);
            Assert.Equal(47625, extent.Cy);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void ImageInsideTableTdWithInheritedWidthAndHeight()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><td style=\"height:3px;width:4px\"><img src=\"data:image/png,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" style=\"height:2px\" /></td></tr></table>"));

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

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(2, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);
            Assert.Equal(1, para.ChildElements.Count);

            DW.Inline inline = image.ChildElements[0] as DW.Inline;
            Assert.NotNull(inline);
            Assert.Equal(5, inline.ChildElements.Count);

            DW.Extent extent = inline.ChildElements[0] as DW.Extent;
            Assert.NotNull(extent);
            Assert.Equal(38100, extent.Cx);
            Assert.Equal(19050, extent.Cy);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void ImageWithInheritedPWidth()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p style=\"width:10px\"><img src=\"data:image/png,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\" alt=\"Red dot\" /></p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Drawing image = run.ChildElements[0] as Drawing;
            Assert.NotNull(image);
            Assert.Equal(1, para.ChildElements.Count);

            DW.Inline inline = image.ChildElements[0] as DW.Inline;
            Assert.NotNull(inline);
            Assert.Equal(5, inline.ChildElements.Count);

            DW.Extent extent = inline.ChildElements[0] as DW.Extent;
            Assert.NotNull(extent);
            Assert.Equal(95250, extent.Cx);
            Assert.Equal(95250, extent.Cy);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }
    }
}
