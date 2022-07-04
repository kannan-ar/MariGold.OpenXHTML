namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using OVML = DocumentFormat.OpenXml.Vml.Office;
    using V = DocumentFormat.OpenXml.Vml;

    public class ObjectElements
    {
        [Fact]
        public void DocxObject()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.BaseURL = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/docxobjecttag.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            EmbeddedObject embeddedObject = run.ChildElements[0] as EmbeddedObject;
            Assert.NotNull(embeddedObject);
            Assert.Equal(2, embeddedObject.ChildElements.Count);

            V.Shape shape = embeddedObject.ChildElements[0] as V.Shape;
            Assert.NotNull(shape);

            OVML.OleObject oleObject = embeddedObject.ChildElements[1] as OVML.OleObject;
            Assert.NotNull(oleObject);

            Assert.Equal(OVML.OleValues.Embed, oleObject.Type.Value);
            Assert.Equal("Word.Document.12", oleObject.ProgId.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void PptxObject()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.BaseURL = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/pptxobjecttag.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            EmbeddedObject embeddedObject = run.ChildElements[0] as EmbeddedObject;
            Assert.NotNull(embeddedObject);
            Assert.Equal(2, embeddedObject.ChildElements.Count);

            V.Shape shape = embeddedObject.ChildElements[0] as V.Shape;
            Assert.NotNull(shape);

            OVML.OleObject oleObject = embeddedObject.ChildElements[1] as OVML.OleObject;
            Assert.NotNull(oleObject);

            Assert.Equal(OVML.OleValues.Embed, oleObject.Type.Value);
            Assert.Equal("PowerPoint.Show.12", oleObject.ProgId.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void XlsxObject()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);
            string path = TestUtility.GetPath("/Html");
            doc.BaseURL = path;
            doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/xlsxobjecttag.htm")));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            EmbeddedObject embeddedObject = run.ChildElements[0] as EmbeddedObject;
            Assert.NotNull(embeddedObject);
            Assert.Equal(2, embeddedObject.ChildElements.Count);

            V.Shape shape = embeddedObject.ChildElements[0] as V.Shape;
            Assert.NotNull(shape);

            OVML.OleObject oleObject = embeddedObject.ChildElements[1] as OVML.OleObject;
            Assert.NotNull(oleObject);

            Assert.Equal(OVML.OleValues.Embed, oleObject.Type.Value);
            Assert.Equal("Excel.Sheet.12", oleObject.ProgId.Value);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }
    }
}
