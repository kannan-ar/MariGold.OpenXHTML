namespace MariGold.OpenXHTML.Tests
{
	using System;
	using NUnit.Framework;
	using MariGold.OpenXHTML;
	using System.IO;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	using Word = DocumentFormat.OpenXml.Wordprocessing;
	using DocumentFormat.OpenXml.Validation;
	using System.Linq;
	
	[TestFixture]
	public class BasicStyles
	{
		[Test]
		public void DivColorRed()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div style='color:#ff0000'>test</div>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				OpenXmlElement para = doc.Document.Body.ChildElements[0];
				
				Assert.IsTrue(para is Paragraph);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				Assert.IsNotNull(run.RunProperties);
				Assert.IsNotNull(run.RunProperties.Color);
				Assert.AreEqual("ff0000", run.RunProperties.Color.Val.Value);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void DivRGBColorRed()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div style='color:rgb(255,0,0)'>test</div>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				OpenXmlElement para = doc.Document.Body.ChildElements[0];
				
				Assert.IsTrue(para is Paragraph);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				Assert.IsNotNull(run.RunProperties);
				Assert.IsNotNull(run.RunProperties.Color);
				Assert.AreEqual("FF0000", run.RunProperties.Color.Val.Value);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void iTag()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<i>test</i>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				OpenXmlElement para = doc.Document.Body.ChildElements[0];
				
				Assert.IsTrue(para is Paragraph);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				Assert.IsNotNull(run.RunProperties);
				Assert.AreEqual(1, run.RunProperties.ChildElements.Count);
				Italic italic = run.RunProperties.ChildElements[0] as Italic;
				Assert.IsNotNull(italic);
				
				Word.Text text = run.ChildElements[1] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual(0, text.ChildElements.Count);
				Assert.AreEqual("test", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void DivUnderline()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div style='text-decoration:underline'>test</div>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				OpenXmlElement para = doc.Document.Body.ChildElements[0];
				
				Assert.IsTrue(para is Paragraph);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				Assert.IsNotNull(run.RunProperties);
				Assert.AreEqual(1, run.RunProperties.ChildElements.Count);
				Underline underline = run.RunProperties.ChildElements[0] as Underline;
				Assert.IsNotNull(underline);
				Assert.AreEqual(UnderlineValues.Single, underline.Val.Value);
				
				Word.Text text = run.ChildElements[1] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual(0, text.ChildElements.Count);
				Assert.AreEqual("test", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
			
		}
		
		[Test]
		public void bTag()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<b>test</b>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				OpenXmlElement para = doc.Document.Body.ChildElements[0];
				
				Assert.IsTrue(para is Paragraph);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				Assert.IsNotNull(run.RunProperties);
				Assert.AreEqual(1, run.RunProperties.ChildElements.Count);
				Bold bold = run.RunProperties.ChildElements[0] as Bold;
				Assert.IsNotNull(bold);
				
				Word.Text text = run.ChildElements[1] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual(0, text.ChildElements.Count);
				Assert.AreEqual("test", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void StrongTagWithSpace()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<strong>Name &amp; SSN </string>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				OpenXmlElement para = doc.Document.Body.ChildElements[0];
				
				Assert.IsTrue(para is Paragraph);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				Assert.IsNotNull(run.RunProperties);
				Assert.AreEqual(1, run.RunProperties.ChildElements.Count);
				Bold bold = run.RunProperties.ChildElements[0] as Bold;
				Assert.IsNotNull(bold);
				
				Word.Text text = run.ChildElements[1] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual(0, text.ChildElements.Count);
				Assert.AreEqual("Name & SSN ", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}

        [Test]
        public void BackgroundProperty()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<div style='background:#000 no-repeat right top;'>test</div>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.AreEqual(2, paragraph.ChildElements.Count);
                Assert.IsNotNull(paragraph.ParagraphProperties);
                Assert.IsNotNull(paragraph.ParagraphProperties.Shading);
                Assert.AreEqual("000000", paragraph.ParagraphProperties.Shading.Fill.Value);
                Assert.AreEqual(Word.ShadingPatternValues.Clear, paragraph.ParagraphProperties.Shading.Val.Value);

                Run run = paragraph.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual(0, text.ChildElements.Count);
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }
	}
}
