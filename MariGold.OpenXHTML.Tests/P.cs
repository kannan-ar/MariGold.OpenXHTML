namespace MariGold.OpenXHTML.Tests
{
	using System;
	using NUnit.Framework;
	using MariGold.OpenXHTML;
	using System.IO;
	using DocumentFormat.OpenXml.Wordprocessing;
	using Word = DocumentFormat.OpenXml.Wordprocessing;
	using DocumentFormat.OpenXml.Validation;
	using System.Linq;
	
	[TestFixture]
	public class P
	{
		[Test]
		public void SinglePBackGround()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<p style='background-color:#000'>test</p>"));
				
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
		
		[Test]
		public void SinglePRedColor()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<p style='color:red'>test</p>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
				
				Run run = paragraph.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				Assert.IsNotNull(run.RunProperties);
				Word.Color color = run.RunProperties.ChildElements[0] as Word.Color;
				Assert.AreEqual("FF0000", color.Val.Value);
				
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
