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
	public class TestFonts
	{
		[Test]
		public void DivFontBold()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div style='font-weight:bold'>test</div>"));
				
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
		public void DivFontFamily()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div style='font-family:arial'>test</div>"));
				
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
				RunFonts fonts = run.RunProperties.ChildElements[0] as RunFonts;
				Assert.IsNotNull(fonts);
				Assert.AreEqual("arial",fonts.Ascii.Value);
				
				Word.Text text = run.ChildElements[1] as Word.Text;
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
