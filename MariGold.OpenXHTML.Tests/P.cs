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
		
		[Test]
		public void SinglePAllBorder()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<p style='border:1px solid #000'>test</p>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(paragraph);
				Assert.AreEqual(2, paragraph.ChildElements.Count);
				
				ParagraphProperties paragraphProperties = paragraph.ChildElements[0] as ParagraphProperties;
				ParagraphBorders paragraphBorders = paragraphProperties.ChildElements[0] as ParagraphBorders;
				Assert.IsNotNull(paragraphBorders);
				Assert.AreEqual(4, paragraphBorders.ChildElements.Count);
				
				TopBorder topBorder = paragraphBorders.ChildElements[0] as TopBorder;
				Assert.IsNotNull(topBorder);
				Assert.AreEqual(BorderValues.Single, topBorder.Val.Value);
				Assert.AreEqual("000000", topBorder.Color.Value);
				Assert.AreEqual(1, topBorder.Size.Value);
				
				LeftBorder leftBorder = paragraphBorders.ChildElements[1] as LeftBorder;
				Assert.IsNotNull(leftBorder);
				Assert.AreEqual(BorderValues.Single, leftBorder.Val.Value);
				Assert.AreEqual("000000", leftBorder.Color.Value);
				Assert.AreEqual(1, leftBorder.Size.Value);
				
				BottomBorder bottomBorder = paragraphBorders.ChildElements[2] as BottomBorder;
				Assert.IsNotNull(bottomBorder);
				Assert.AreEqual(BorderValues.Single, bottomBorder.Val.Value);
				Assert.AreEqual("000000", bottomBorder.Color.Value);
				Assert.AreEqual(1, bottomBorder.Size.Value);
				
				RightBorder rightBorder = paragraphBorders.ChildElements[3] as RightBorder;
				Assert.IsNotNull(rightBorder);
				Assert.AreEqual(BorderValues.Single, rightBorder.Val.Value);
				Assert.AreEqual("000000", rightBorder.Color.Value);
				Assert.AreEqual(1, rightBorder.Size.Value);
				
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
		public void SinglePAllBorderWithBorderBottom()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<p style='border:1px solid #000;border-bottom:red solid 2px'>test</p>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(paragraph);
				Assert.AreEqual(2, paragraph.ChildElements.Count);
				
				ParagraphProperties paragraphProperties = paragraph.ChildElements[0] as ParagraphProperties;
				ParagraphBorders paragraphBorders = paragraphProperties.ChildElements[0] as ParagraphBorders;
				Assert.IsNotNull(paragraphBorders);
				Assert.AreEqual(4, paragraphBorders.ChildElements.Count);
				
				TopBorder topBorder = paragraphBorders.ChildElements[0] as TopBorder;
				Assert.IsNotNull(topBorder);
				Assert.AreEqual(BorderValues.Single, topBorder.Val.Value);
				Assert.AreEqual("000000", topBorder.Color.Value);
				Assert.AreEqual(1, topBorder.Size.Value);
				
				LeftBorder leftBorder = paragraphBorders.ChildElements[1] as LeftBorder;
				Assert.IsNotNull(leftBorder);
				Assert.AreEqual(BorderValues.Single, leftBorder.Val.Value);
				Assert.AreEqual("000000", leftBorder.Color.Value);
				Assert.AreEqual(1, leftBorder.Size.Value);
				
				BottomBorder bottomBorder = paragraphBorders.ChildElements[2] as BottomBorder;
				Assert.IsNotNull(bottomBorder);
				Assert.AreEqual(BorderValues.Single, bottomBorder.Val.Value);
				Assert.AreEqual("FF0000", bottomBorder.Color.Value);
				Assert.AreEqual(2, bottomBorder.Size.Value);
				
				RightBorder rightBorder = paragraphBorders.ChildElements[3] as RightBorder;
				Assert.IsNotNull(rightBorder);
				Assert.AreEqual(BorderValues.Single, rightBorder.Val.Value);
				Assert.AreEqual("000000", rightBorder.Color.Value);
				Assert.AreEqual(1, rightBorder.Size.Value);
				
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
		public void SinglePAllBorderWithIndependentBorders()
		{
			using (MemoryStream mem = new MemoryStream())
			{
				WordDocument doc = new WordDocument(mem);
			
				doc.Process(new HtmlParser("<p style='border:1px solid #000;border-bottom:red solid 2px;border-top:blue 3px solid;border-left:#F0F000 4px solid;border-right:solid #CCC888 5px'>test</p>"));
				
				Assert.IsNotNull(doc.Document.Body);
				Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);
				
				Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
				Assert.IsNotNull(paragraph);
				Assert.AreEqual(2, paragraph.ChildElements.Count);
				
				ParagraphProperties paragraphProperties = paragraph.ChildElements[0] as ParagraphProperties;
				ParagraphBorders paragraphBorders = paragraphProperties.ChildElements[0] as ParagraphBorders;
				Assert.IsNotNull(paragraphBorders);
				Assert.AreEqual(4, paragraphBorders.ChildElements.Count);
				
				TopBorder topBorder = paragraphBorders.ChildElements[0] as TopBorder;
				Assert.IsNotNull(topBorder);
				Assert.AreEqual(BorderValues.Single, topBorder.Val.Value);
				Assert.AreEqual("0000FF", topBorder.Color.Value);
				Assert.AreEqual(3, topBorder.Size.Value);
				
				LeftBorder leftBorder = paragraphBorders.ChildElements[1] as LeftBorder;
				Assert.IsNotNull(leftBorder);
				Assert.AreEqual(BorderValues.Single, leftBorder.Val.Value);
				Assert.AreEqual("f0f000", leftBorder.Color.Value);
				Assert.AreEqual(4, leftBorder.Size.Value);
				
				BottomBorder bottomBorder = paragraphBorders.ChildElements[2] as BottomBorder;
				Assert.IsNotNull(bottomBorder);
				Assert.AreEqual(BorderValues.Single, bottomBorder.Val.Value);
				Assert.AreEqual("FF0000", bottomBorder.Color.Value);
				Assert.AreEqual(2, bottomBorder.Size.Value);
				
				RightBorder rightBorder = paragraphBorders.ChildElements[3] as RightBorder;
				Assert.IsNotNull(rightBorder);
				Assert.AreEqual(BorderValues.Single, rightBorder.Val.Value);
				Assert.AreEqual("ccc888", rightBorder.Color.Value);
				Assert.AreEqual(5, rightBorder.Size.Value);
				
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
