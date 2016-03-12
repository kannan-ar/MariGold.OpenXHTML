namespace MariGold.OpenXHTML.Tests
{
	using System;
	using NUnit.Framework;
	using MariGold.OpenXHTML;
	using DocumentFormat.OpenXml.Wordprocessing;
	using Word = DocumentFormat.OpenXml.Wordprocessing;
	using DocumentFormat.OpenXml.Validation;
	using System.IO;
	using System.Linq;
	
	[TestFixture]
	public class SimpleHtmlFromFile
	{
		[Test]
		public void EmptyHtmlBody()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser(TestUtility.GetHtmlFromFile("Html\\emptybody.htm")));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0]as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				
				Word.Text text = run.ChildElements[0]as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual(0, text.ChildElements.Count);
				Assert.AreEqual(string.Empty, text.InnerText.Trim());
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				errors.PrintValidationErrors();
				Assert.AreEqual(0, errors.Count());
			}
		}
	}
}
