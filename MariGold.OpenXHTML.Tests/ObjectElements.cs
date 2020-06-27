namespace MariGold.OpenXHTML.Tests
{
    using NUnit.Framework;
    using OpenXHTML;
    using System.IO;
    using DocumentFormat.OpenXml.Wordprocessing;
    using V = DocumentFormat.OpenXml.Vml;
    using OVML = DocumentFormat.OpenXml.Vml.Office;
    using DocumentFormat.OpenXml.Validation;
    using System.Linq;

    [TestFixture]
    public class ObjectElements
    {
        [Test]
        public void DocxObject()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);
                string path = TestUtility.GetPath("/Html");
                doc.BaseURL = path;
                doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/docxobjecttag.htm")));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(paragraph);

                Run run = paragraph.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                EmbeddedObject embeddedObject = run.ChildElements[0] as EmbeddedObject;
                Assert.IsNotNull(embeddedObject);
                Assert.AreEqual(2, embeddedObject.ChildElements.Count);

                V.Shape shape = embeddedObject.ChildElements[0] as V.Shape;
                Assert.IsNotNull(shape);

                OVML.OleObject oleObject = embeddedObject.ChildElements[1] as OVML.OleObject;
                Assert.IsNotNull(oleObject);

                Assert.AreEqual(OVML.OleValues.Embed, oleObject.Type.Value);
                Assert.AreEqual("Word.Document.12", oleObject.ProgId.Value);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                errors.PrintValidationErrors();
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void PptxObject()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);
                string path = TestUtility.GetPath("/Html");
                doc.BaseURL = path;
                doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/pptxobjecttag.htm")));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(paragraph);

                Run run = paragraph.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                EmbeddedObject embeddedObject = run.ChildElements[0] as EmbeddedObject;
                Assert.IsNotNull(embeddedObject);
                Assert.AreEqual(2, embeddedObject.ChildElements.Count);

                V.Shape shape = embeddedObject.ChildElements[0] as V.Shape;
                Assert.IsNotNull(shape);

                OVML.OleObject oleObject = embeddedObject.ChildElements[1] as OVML.OleObject;
                Assert.IsNotNull(oleObject);

                Assert.AreEqual(OVML.OleValues.Embed, oleObject.Type.Value);
                Assert.AreEqual("PowerPoint.Show.12", oleObject.ProgId.Value);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                errors.PrintValidationErrors();
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void XlsxObject()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);
                string path = TestUtility.GetPath("/Html");
                doc.BaseURL = path;
                doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("/Html/xlsxobjecttag.htm")));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(paragraph);

                Run run = paragraph.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                EmbeddedObject embeddedObject = run.ChildElements[0] as EmbeddedObject;
                Assert.IsNotNull(embeddedObject);
                Assert.AreEqual(2, embeddedObject.ChildElements.Count);

                V.Shape shape = embeddedObject.ChildElements[0] as V.Shape;
                Assert.IsNotNull(shape);

                OVML.OleObject oleObject = embeddedObject.ChildElements[1] as OVML.OleObject;
                Assert.IsNotNull(oleObject);

                Assert.AreEqual(OVML.OleValues.Embed, oleObject.Type.Value);
                Assert.AreEqual("Excel.Sheet.12", oleObject.ProgId.Value);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                errors.PrintValidationErrors();
                Assert.AreEqual(0, errors.Count());
            }
        }
    }
}
