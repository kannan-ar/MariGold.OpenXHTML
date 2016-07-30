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
    public class HtmlDefaultStyles
    {
        [Test]
        public void ThWithSpanAndDiv()
        {
            using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);

                doc.Process(new HtmlParser("<table><tr><th><span>one</span><div>two</div></th></tr></table>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;
				
				Assert.IsNotNull(table);
				Assert.AreEqual(3, table.ChildElements.Count);
				
				TableProperties tableProperties = table.ChildElements[0] as TableProperties;
				Assert.IsNotNull(tableProperties);
				
				TableStyle tableStyle = tableProperties.ChildElements[0]as TableStyle;
				Assert.IsNotNull(tableStyle);
				Assert.AreEqual("TableGrid", tableStyle.Val.Value);
				
				TableGrid tableGrid = table.ChildElements[1] as TableGrid;
				Assert.IsNotNull(tableGrid);
				Assert.AreEqual(1, tableGrid.ChildElements.Count);
				
				TableRow row = table.ChildElements[2] as TableRow;
				Assert.IsNotNull(row);
				Assert.AreEqual(1, row.ChildElements.Count);
				
				TableCell cell = row.ChildElements[0] as TableCell;
				Assert.IsNotNull(cell);
				Assert.AreEqual(2, cell.ChildElements.Count);
				
				Paragraph para = cell.ChildElements[0] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                Run run = para.ChildElements[0] as Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);
                Assert.IsNotNull(run.RunProperties);
                Bold bold = run.RunProperties.ChildElements[0] as Bold;
                Assert.IsNotNull(bold);

                Word.Text text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual(0, text.ChildElements.Count);
                Assert.AreEqual("one", text.InnerText);

                para = cell.ChildElements[1] as Paragraph;
                Assert.IsNotNull(para);
                Assert.AreEqual(1, para.ChildElements.Count);

                run = para.ChildElements[0] as Run;

                Assert.IsNotNull(run);
                Assert.AreEqual(2, run.ChildElements.Count);
                Assert.IsNotNull(run.RunProperties);
                bold = run.RunProperties.ChildElements[0] as Bold;
                Assert.IsNotNull(bold);

                text = run.ChildElements[1] as Word.Text;
                Assert.IsNotNull(text);
                Assert.AreEqual(0, text.ChildElements.Count);
                Assert.AreEqual("two", text.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc.WordprocessingDocument);
                Assert.AreEqual(0, errors.Count());
            }
        }
    }
}
