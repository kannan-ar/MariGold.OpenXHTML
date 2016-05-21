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
				Assert.AreEqual(0, doc.Document.Body.ChildElements.Count);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
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
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
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
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
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
				
				Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;
				
				Assert.IsNotNull(table);
				Assert.AreEqual(3, table.ChildElements.Count);
				
				TableRow row = table.ChildElements[2] as TableRow;
				
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
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
				
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
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
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
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
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
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
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
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
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
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
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
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
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
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void TextAndSpanOnDiv()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div>pp<span>test1</span></div>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				
				Assert.IsNotNull(para);
				Assert.AreEqual(2, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				var text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("pp", text.InnerText);
				
				run = para.ChildElements[1] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("test1", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void DivInsideAnotherDiv()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div><div>test</div></div>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				var text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("test", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void TextWithBreak()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("test<br />text"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(3, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				var text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("test", text.InnerText);
				
				run = para.ChildElements[1] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				var br = run.ChildElements[0] as Break;
				Assert.IsNotNull(br);
				
				run = para.ChildElements[2] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("text", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void DivTextWithBreak()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div>test<br />text</div>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(3, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				var text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("test", text.InnerText);
				
				run = para.ChildElements[1] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				var br = run.ChildElements[0] as Break;
				Assert.IsNotNull(br);
				
				run = para.ChildElements[2] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("text", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void DivSpanTextWithBreak()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div><span>test</span><br /><span>text</span></div>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(3, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				var text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("test", text.InnerText);
				
				run = para.ChildElements[1] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				var br = run.ChildElements[0] as Break;
				Assert.IsNotNull(br);
				
				run = para.ChildElements[2] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("text", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void DivSpanStyleTextWithBreak()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div><span style='color:#ff0000'>test</span><br /><span>text</span></div>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(3, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				var text = run.ChildElements[1] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("test", text.InnerText);
				
				Assert.IsNotNull(run.RunProperties);
				Assert.IsNotNull(run.RunProperties.Color);
				Assert.AreEqual("ff0000", run.RunProperties.Color.Val.Value);
				
				run = para.ChildElements[1] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				var br = run.ChildElements[0] as Break;
				Assert.IsNotNull(br);
				
				run = para.ChildElements[2] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("text", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void InnerDivAndText()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<div><div>one</div>two</div>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(2, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				var text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("one", text.InnerText);
				
				para = doc.Document.Body.ChildElements[1] as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("two", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void H1Only()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<h1>test</h1>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				var text = run.ChildElements[1] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("test", text.InnerText);
				
				Assert.IsNotNull(run.RunProperties);
				Bold bold = run.RunProperties.ChildElements[0] as Bold;
				Assert.IsNotNull(bold);
				FontSize fontSize = run.RunProperties.ChildElements[1] as FontSize;
				Assert.IsNotNull(fontSize);
				Assert.AreEqual("32", fontSize.Val.Value);
				
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void SimpleAddress()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<address>first line<br />second line</address>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				OpenXmlElement para = doc.Document.Body.ChildElements[0];
				
				Assert.IsTrue(para is Paragraph);
				Assert.AreEqual(3, para.ChildElements.Count);
				
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
				Assert.AreEqual("first line", text.InnerText);
				
				run = para.ChildElements[1] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				Break br = run.ChildElements[1] as Break;
				Assert.IsNotNull(br);
				
				run = para.ChildElements[2] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				Assert.IsNotNull(run.RunProperties);
				Assert.AreEqual(1, run.RunProperties.ChildElements.Count);
				italic = run.RunProperties.ChildElements[0] as Italic;
				Assert.IsNotNull(italic);
				
				text = run.ChildElements[1] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual(0, text.ChildElements.Count);
				Assert.AreEqual("second line", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void SimpleDL()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<dl><dt>Numbers</dt><dd>1</dd><dt>Text</dt><dd>One</dd></dl>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(6, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.AreEqual(1, para.ChildElements.Count);
				ParagraphProperties properties = para.ChildElements[0]as ParagraphProperties;
				Assert.IsNotNull(properties);
				Assert.AreEqual(1, properties.ChildElements.Count);
				SpacingBetweenLines spacing = properties.ChildElements[0] as SpacingBetweenLines;
				Assert.IsNotNull(spacing);
				Assert.AreEqual("240", spacing.Before.Value);
				
				para = doc.Document.Body.ChildElements[1] as Paragraph;
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Run run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				Word.Text text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("Numbers", text.InnerText);
				
				para = doc.Document.Body.ChildElements[2] as Paragraph;
				Assert.AreEqual(2, para.ChildElements.Count);
				properties = para.ChildElements[0]as ParagraphProperties;
				Assert.IsNotNull(properties);
				Assert.AreEqual(1, properties.ChildElements.Count);
				Indentation ind = properties.ChildElements[0] as Indentation;
				Assert.IsNotNull(ind);
				Assert.AreEqual("800", ind.Left.Value);
				Assert.IsNull(ind.Right);
				
				run = para.ChildElements[1] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("1", text.InnerText);
				
				para = doc.Document.Body.ChildElements[3] as Paragraph;
				Assert.AreEqual(1, para.ChildElements.Count);
				
				run = para.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("Text", text.InnerText);
				
				para = doc.Document.Body.ChildElements[4] as Paragraph;
				Assert.AreEqual(2, para.ChildElements.Count);
				properties = para.ChildElements[0]as ParagraphProperties;
				Assert.IsNotNull(properties);
				Assert.AreEqual(1, properties.ChildElements.Count);
				ind = properties.ChildElements[0] as Indentation;
				Assert.IsNotNull(ind);
				Assert.AreEqual("800", ind.Left.Value);
				Assert.IsNull(ind.Right);
				
				run = para.ChildElements[1] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("One", text.InnerText);
				
				para = doc.Document.Body.ChildElements[5] as Paragraph;
				Assert.AreEqual(1, para.ChildElements.Count);
				properties = para.ChildElements[0]as ParagraphProperties;
				Assert.IsNotNull(properties);
				spacing = properties.ChildElements[0] as SpacingBetweenLines;
				Assert.IsNotNull(spacing);
				Assert.IsNull(spacing.Before);
				Assert.AreEqual("240", spacing.After.Value);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				errors.PrintValidationErrors();
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void OnlyHr()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<hr />"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.AreEqual(2, para.ChildElements.Count);
				ParagraphProperties properties = para.ChildElements[0]as ParagraphProperties;
				Assert.IsNotNull(properties);
				Assert.AreEqual(1, properties.ChildElements.Count);
				
				ParagraphBorders borders = properties.ChildElements[0] as ParagraphBorders;
				Assert.IsNotNull(borders);
				Assert.AreEqual(1, borders.ChildElements.Count);
				
				TopBorder topBorder = borders.ChildElements[0] as TopBorder;
				Assert.IsNotNull(topBorder);
				TestUtility.TestBorder<TopBorder>(topBorder, BorderValues.Single, "auto", 4U);
				
				Run run = para.ChildElements[1] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(1, run.ChildElements.Count);
				Word.Text text = run.ChildElements[0] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual(string.Empty, text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				errors.PrintValidationErrors();
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void ATag()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<a href='http://google.com'>test</a>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Hyperlink hyperLink = para.ChildElements[0] as Hyperlink;
				Assert.IsNotNull(hyperLink);
				
				Run run = hyperLink.ChildElements[0]as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				RunProperties properties = run.ChildElements[0]as RunProperties;
				Assert.IsNotNull(properties);
				Assert.AreEqual(1, properties.ChildElements.Count);
				
				RunStyle runStyle = properties.ChildElements[0] as RunStyle;
				Assert.IsNotNull(runStyle);
				Assert.AreEqual("Hyperlink", runStyle.Val.Value);
				
				Word.Text text = run.ChildElements[1] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("test", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				errors.PrintValidationErrors();
				Assert.AreEqual(0, errors.Count());
			}
		}
		
		[Test]
		public void ATagWithBold()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<a href='http://google.com'><strong>bold</strong>test</a>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(para);
				Assert.AreEqual(1, para.ChildElements.Count);
				
				Hyperlink hyperLink = para.ChildElements[0] as Hyperlink;
				Assert.IsNotNull(hyperLink);
				
				Run run = hyperLink.ChildElements[0] as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				RunProperties properties = run.ChildElements[0] as RunProperties;
				Assert.IsNotNull(properties);
				Assert.AreEqual(1, properties.ChildElements.Count);
				
				Bold bold = properties.ChildElements[0]as Bold;
				Assert.IsNotNull(bold);
				
				Word.Text text = run.ChildElements[1]as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("bold", text.InnerText);
				
				run = hyperLink.ChildElements[1]as Run;
				Assert.IsNotNull(run);
				Assert.AreEqual(2, run.ChildElements.Count);
				
				properties = run.ChildElements[0]as RunProperties;
				Assert.IsNotNull(properties);
				Assert.AreEqual(1, properties.ChildElements.Count);
				
				RunStyle runStyle = properties.ChildElements[0] as RunStyle;
				Assert.IsNotNull(runStyle);
				Assert.AreEqual("Hyperlink", runStyle.Val.Value);
				
				text = run.ChildElements[1] as Word.Text;
				Assert.IsNotNull(text);
				Assert.AreEqual("test", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				errors.PrintValidationErrors();
				Assert.AreEqual(0, errors.Count());
			}
		}

        [Test]
        public void SpanInsideATag()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<a href='http://google.com'><span>click</span> here</a>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Hyperlink link = para.ChildElements[0] as Hyperlink;

                Assert.IsNotNull(link);
                Assert.AreEqual(2, link.ChildElements.Count);

                Run run = link.ChildElements[0] as Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual("click", text.InnerText);

                run = link.ChildElements[1] as Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;

                Assert.IsNotNull(properties);
                Assert.AreEqual(1, properties.ChildElements.Count);

                RunStyle runStyle = properties.ChildElements[0] as RunStyle;

                Assert.IsNotNull(runStyle);
                Assert.AreEqual("Hyperlink", runStyle.Val.Value);

                text = run.ChildElements[1] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual(" here", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void InsideATagTextSpan()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<a href='http://google.com'>here<span>click</span></a>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Hyperlink link = para.ChildElements[0] as Hyperlink;

                Assert.IsNotNull(link);
                Assert.AreEqual(2, link.ChildElements.Count);

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
                Assert.AreEqual("here", text.InnerText);

                run = link.ChildElements[1] as Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual("click", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void InsideATagTextBr()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<a href='http://google.com'>here<br /></a>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Hyperlink link = para.ChildElements[0] as Hyperlink;

                Assert.IsNotNull(link);
                Assert.AreEqual(2, link.ChildElements.Count);

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
                Assert.AreEqual("here", text.InnerText);

                run = link.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);
                Break br = run.ChildElements[0] as Break;
                Assert.IsNotNull(br);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void InsideATagCenterText()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<a href='http://google.com'><center>click </center>here</a>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Hyperlink link = para.ChildElements[0] as Hyperlink;

                Assert.IsNotNull(link);
                Assert.AreEqual(2, link.ChildElements.Count);

                Run run = link.ChildElements[0] as Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual("click ", text.InnerText);

                run = link.ChildElements[1] as Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;

                Assert.IsNotNull(properties);
                Assert.AreEqual(1, properties.ChildElements.Count);

                RunStyle runStyle = properties.ChildElements[0] as RunStyle;

                Assert.IsNotNull(runStyle);
                Assert.AreEqual("Hyperlink", runStyle.Val.Value);

                text = run.ChildElements[1] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual("here", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void InsideATagItalicText()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<a href='http://google.com'><i>click </i>here</a>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Hyperlink link = para.ChildElements[0] as Hyperlink;

                Assert.IsNotNull(link);
                Assert.AreEqual(2, link.ChildElements.Count);

                Run run = link.ChildElements[0] as Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual("click ", text.InnerText);

                run = link.ChildElements[1] as Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);

                RunProperties properties = run.ChildElements[0] as RunProperties;

                Assert.IsNotNull(properties);
                Assert.AreEqual(1, properties.ChildElements.Count);

                RunStyle runStyle = properties.ChildElements[0] as RunStyle;

                Assert.IsNotNull(runStyle);
                Assert.AreEqual("Hyperlink", runStyle.Val.Value);

                text = run.ChildElements[1] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual("here", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void TwoATagWithSpan()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<a href='http://google.com'><span>one</span></a><a href='#'><span>two</span></a>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;

                Assert.IsNotNull(para);
                Assert.AreEqual(2, para.ChildElements.Count);

                Hyperlink link = para.ChildElements[0] as Hyperlink;

                Assert.IsNotNull(link);
                Assert.AreEqual(1, link.ChildElements.Count);

                Run run = link.ChildElements[0] as Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                Word.Text text = run.ChildElements[0] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual("one", text.InnerText);

                run = para.ChildElements[1] as Run;
                Assert.IsNotNull(run);
                Assert.AreEqual(1, run.ChildElements.Count);

                text = run.ChildElements[0] as Word.Text;

                Assert.IsNotNull(text);
                Assert.AreEqual("two", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                errors.PrintValidationErrors();
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void ProtocolFreeUrl()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);
                doc.UriSchema = Uri.UriSchemeHttp;
                doc.Process(new HtmlParser("<a href='//google.com'>test</a>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

                Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Hyperlink hyperLink = para.ChildElements[0] as Hyperlink;
                Assert.IsNotNull(hyperLink);

                Run run = hyperLink.ChildElements[0] as Run;
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
                Assert.AreEqual("test", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                errors.PrintValidationErrors();
                Assert.AreEqual(0, errors.Count());
            }
        }

        [Test]
        public void DisplayNone()
        {
            using (MemoryStream mem = new MemoryStream())
            {
                WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<div style='display:none'>test</div>"));

                Assert.IsNotNull(doc.Document.Body);
                Assert.AreEqual(0, doc.Document.Body.ChildElements.Count);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                errors.PrintValidationErrors();
                Assert.AreEqual(0, errors.Count());
            }
        }
	}
}
