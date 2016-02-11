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
	public class Tables
	{
		[Test]
		public void TableBorder()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<table border='1'><tr><td>test</td></tr></table>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Table table = doc.Document.Body.ChildElements[0] as Table;
				
				Assert.IsNotNull(table);
				Assert.AreEqual(3, table.ChildElements.Count);
				
				TableProperties tableProperties = table.ChildElements[0] as TableProperties;
				Assert.IsNotNull(tableProperties);
				
				TableStyle tableStyle = tableProperties.ChildElements[0]as TableStyle;
				Assert.IsNotNull(tableStyle);
				Assert.AreEqual("TableGrid", tableStyle.Val.Value);
				
				TableBorders tableBorders = tableProperties.ChildElements[1] as TableBorders;
				Assert.IsNotNull(tableBorders);
				
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
				Assert.AreEqual("test", text.InnerText);
				
				OpenXmlValidator validator = new OpenXmlValidator();
				var errors = validator.Validate(doc.WordprocessingDocument);
				Assert.AreEqual(0, errors.Count());
			}
		}
	}
}
