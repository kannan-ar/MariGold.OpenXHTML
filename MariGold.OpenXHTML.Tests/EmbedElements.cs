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
    public class EmbedElements
    {
        [Test]
        public void SimpleObjectWithImage()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);
                string path = "file:///" + TestUtility.GetPath("Html");
                path = path.Replace(@"\", "//");
                doc.ImagePath = path + "//";
                doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("Html\\objectimage.htm")));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Drawing image = run.ChildElements[0] as Drawing;
                Assert.IsNotNull(image);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                errors.PrintValidationErrors();
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void ImageInsideDiv()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);
                string path = "file:///" + TestUtility.GetPath("Html");
                path = path.Replace(@"\", "//");
                doc.ImagePath = path + "//";
                doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("Html\\objectimageinsidediv.htm")));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Drawing image = run.ChildElements[0] as Drawing;
                Assert.IsNotNull(image);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                errors.PrintValidationErrors();
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void ImageParagraphDiv()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);
                string path = "file:///" + TestUtility.GetPath("Html");
                path = path.Replace(@"\", "//");
                doc.ImagePath = path + "//";
                doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("Html\\objectparagraphimage.htm")));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(2, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                para = doc.Document.Body.ChildElements[1] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Drawing image = run.ChildElements[0] as Drawing;
                Assert.IsNotNull(image);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                errors.PrintValidationErrors();
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void ImageMultiDiv()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);
                string path = "file:///" + TestUtility.GetPath("Html");
                path = path.Replace(@"\", "//");
                doc.ImagePath = path + "//";
                doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("Html\\objectimagemultidiv.htm")));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(3, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test", text.InnerText);

                para = doc.Document.Body.ChildElements[1] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Drawing image = run.ChildElements[0] as Drawing;
                Assert.IsNotNull(image);

                para = doc.Document.Body.ChildElements[2] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                run = para.ChildElements[0] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual("test1", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                errors.PrintValidationErrors();
                Assert.AreEqual(0, errors.Count());
            }
        }
    }
}
