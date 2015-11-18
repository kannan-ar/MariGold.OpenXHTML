namespace MariGold.OpenXHTML.Tests
{
	using System;
	using NUnit.Framework;
	using MariGold.OpenXHTML;
	using System.IO;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	[TestFixture]
	public class BasicDocument
	{
		[Test]
		public void EmptyString()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser(" "));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				OpenXmlElement para = doc.Document.Body.ChildElements[0];
				
				Assert.IsTrue(para is Paragraph);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				OpenXmlElement run = para.ChildElements[0];
				Assert.IsTrue(run is Run);
				Assert.AreEqual(1, run.ChildElements.Count);
				
				OpenXmlElement text = run.ChildElements[0] as DocumentFormat.OpenXml.Wordprocessing.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual(0, text.ChildElements.Count);
				Assert.AreEqual(" ", text.InnerText);
			}
			
		}
		
		[Test]
		public void EmptyBody()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<body></body>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(0, doc.Document.Body.ChildElements.Count);
			}
		}
		
		[Test]
		public void SimpleText()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("test"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				OpenXmlElement para = doc.Document.Body.ChildElements[0];
				
				Assert.IsTrue(para is Paragraph);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				OpenXmlElement run = para.ChildElements[0];
				Assert.IsTrue(run is Run);
				Assert.AreEqual(1, run.ChildElements.Count);
				
				OpenXmlElement text = run.ChildElements[0] as DocumentFormat.OpenXml.Wordprocessing.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual(0, text.ChildElements.Count);
				Assert.AreEqual("test", text.InnerText);
			}
		}
	}
}
