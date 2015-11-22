namespace MariGold.OpenXHTML.Tests
{
	using System;
	using NUnit.Framework;
	using MariGold.OpenXHTML;
	using System.IO;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	using Word = DocumentFormat.OpenXml.Wordprocessing;
	
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
		
		[Test]
		public void SimpleTable()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<table><tr><td>1</td></tr></table>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Table table = doc.Document.Body.ChildElements[0] as Table;
				
				Assert.IsNotNull(table);
				Assert.AreEqual(1, table.ChildElements.Count);
				
				TableRow row = table.ChildElements[0] as TableRow;
				
				Assert.IsNotNull(row);
				Assert.AreEqual(1, row.ChildElements.Count);
				
				TableCell cell = row.ChildElements[0] as TableCell;
				
				Assert.IsNotNull(cell);
				Assert.AreEqual(1, cell.ChildElements.Count);
				
				Paragraph para = cell.ChildElements[0] as Paragraph;
				
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				
				Word.Text text = run.ChildElements[0] as Word.Text;
				
				Assert.IsNotNull(text);
				Assert.AreEqual(0, text.ChildElements.Count);
				Assert.AreEqual("1", text.InnerText);
			}
		}
		
		[Test]
		public void TwoSpanOnBody()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<span>1</span><span>2</span>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				
				Assert.IsNotNull(para);
				Assert.AreEqual(2, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				
				Word.Text text = run.ChildElements[0] as Word.Text;
				
				Assert.IsNotNull(text);
				Assert.AreEqual("1", text.InnerText);
				
				run = para.ChildElements[1] as Run;
				
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				
				text = run.ChildElements[0] as Word.Text;
				
				Assert.IsNotNull(text);
				Assert.AreEqual("2", text.InnerText);
			}
		}
				
		[Test]
		public void DivAndSpanOnly()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div>1</div><span>2</span>"));
				
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
				Assert.AreEqual("1", text.InnerText);
				
				para = doc.Document.Body.ChildElements[1] as Paragraph;
				
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				run = para.ChildElements[0] as Run;
				
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				
				text = run.ChildElements[0] as Word.Text;
				
				Assert.IsNotNull(text);
				Assert.AreEqual("2", text.InnerText);
			}
		}
		
				
		[Test]
		public void SpanAndDivOnly()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<span>1</span><div>2</div>"));
				
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
				Assert.AreEqual("1", text.InnerText);
				
				para = doc.Document.Body.ChildElements[1] as Paragraph;
				
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				run = para.ChildElements[0] as Run;
				
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				
				text = run.ChildElements[0] as Word.Text;
				
				Assert.IsNotNull(text);
				Assert.AreEqual("2", text.InnerText);
			}
		}
		
		[Test]
		public void DivAndDivOnly()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div>1</div><div>2</div>"));
				
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
				Assert.AreEqual("1", text.InnerText);
				
				para = doc.Document.Body.ChildElements[1] as Paragraph;
				
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				run = para.ChildElements[0] as Run;
				
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				
				text = run.ChildElements[0] as Word.Text;
				
				Assert.IsNotNull(text);
				Assert.AreEqual("2", text.InnerText);
			}
		}
		
		[Test]
		public void DivTwoSpanOnBody()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div>1</div><span>2</span><span>3</span>"));
				
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
				Assert.AreEqual("1", text.InnerText);
				
				para = doc.Document.Body.ChildElements[1] as Paragraph;
				
				Assert.IsNotNull(para);
				Assert.AreEqual(2, para.ChildElements.Count);
				
				run = para.ChildElements[0] as Run;
				text = run.ChildElements[0] as Word.Text;
				Assert.AreEqual("2", text.InnerText);
				
				run = para.ChildElements[1] as Run;
				text = run.ChildElements[0] as Word.Text;
				Assert.AreEqual("3", text.InnerText);
			}
		}
		
		[Test]
		public void OneAOnBody()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<a href='http://google.com'>click here</a>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Hyperlink link = para.ChildElements[0] as Hyperlink;
				
				Assert.IsNotNull(link);
				Assert.AreEqual(1, link.ChildElements.Count);
				
				Run run = link.ChildElements[0] as Run;
				
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				RunProperties properties = run.ChildElements[0] as RunProperties;
				
				Assert.IsNotNull(properties);
				Assert.AreEqual(1, properties.ChildElements.Count);
				
				RunStyle runStyle = properties.ChildElements[0] as RunStyle;
				
				Assert.IsNotNull(runStyle);
				Assert.AreEqual("Hyperlink", runStyle.Val.Value);
				
				Word.Text text = run.ChildElements[1] as Word.Text;
				
				Assert.IsNotNull(text);
				Assert.AreEqual("click here", text.InnerText);
			}
		}
		
		[Test]
		public void AOnDivBody()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div><a href='http://google.com'>click here</a></div>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Hyperlink link = para.ChildElements[0] as Hyperlink;
				
				Assert.IsNotNull(link);
				Assert.AreEqual(1, link.ChildElements.Count);
				
				Run run = link.ChildElements[0] as Run;
				
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				RunProperties properties = run.ChildElements[0] as RunProperties;
				
				Assert.IsNotNull(properties);
				Assert.AreEqual(1, properties.ChildElements.Count);
				
				RunStyle runStyle = properties.ChildElements[0] as RunStyle;
				
				Assert.IsNotNull(runStyle);
				Assert.AreEqual("Hyperlink", runStyle.Val.Value);
				
				Word.Text text = run.ChildElements[1] as Word.Text;
				
				Assert.IsNotNull(text);
				Assert.AreEqual("click here", text.InnerText);
			}
		}
	}
}
