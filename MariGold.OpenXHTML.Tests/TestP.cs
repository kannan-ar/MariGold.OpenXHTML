namespace MariGold.OpenXHTML.Tests
{
	using NUnit.Framework;
	using OpenXHTML;
	using System.IO;
	using DocumentFormat.OpenXml.Wordprocessing;
	using Word = DocumentFormat.OpenXml.Wordprocessing;
	using DocumentFormat.OpenXml.Validation;
	using System.Linq;
	
	[TestFixture]
	public class TestP
	{
		[Test]
		public void SinglePBackGround()
		{
			using MemoryStream mem = new MemoryStream();
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
			Assert.AreEqual(2, run.ChildElements.Count);

			RunProperties runProperties = run.ChildElements[0] as RunProperties;
			Assert.IsNotNull(runProperties);
			Assert.AreEqual("000000", runProperties.Shading.Fill.Value);

			Word.Text text = run.ChildElements[1] as Word.Text;
			Assert.IsNotNull(text);
			Assert.AreEqual(0, text.ChildElements.Count);
			Assert.AreEqual("test", text.InnerText);

			OpenXmlValidator validator = new OpenXmlValidator();
			var errors = validator.Validate(doc.WordprocessingDocument);
			Assert.AreEqual(0, errors.Count());
		}
		
		[Test]
		public void SinglePRedColor()
		{
			using MemoryStream mem = new MemoryStream();
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
		
		[Test]
		public void SinglePAllBorder()
		{
			using MemoryStream mem = new MemoryStream();
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
		
		[Test]
		public void SinglePAllBorderWithBorderBottom()
		{
			using MemoryStream mem = new MemoryStream();
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
		
		[Test]
		public void SinglePAllBorderWithIndependentBorders()
		{
			using MemoryStream mem = new MemoryStream();
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
		
		[Test]
		public void SinglePFontSize()
		{
			using MemoryStream mem = new MemoryStream();
			WordDocument doc = new WordDocument(mem);

			doc.Process(new HtmlParser("<p style='font-size:10px'>test</p>"));

			Assert.IsNotNull(doc.Document.Body);
			Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

			Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;

			Run run = paragraph.ChildElements[0] as Run;
			Assert.IsNotNull(run);
			Assert.AreEqual(2, run.ChildElements.Count);
			Assert.IsNotNull(run.RunProperties);
			FontSize fontSize = run.RunProperties.ChildElements[0] as FontSize;
			Assert.AreEqual("20", fontSize.Val.Value);

			Word.Text text = run.ChildElements[1] as Word.Text;
			Assert.IsNotNull(text);
			Assert.AreEqual(0, text.ChildElements.Count);
			Assert.AreEqual("test", text.InnerText);

			OpenXmlValidator validator = new OpenXmlValidator();
			var errors = validator.Validate(doc.WordprocessingDocument);
			Assert.AreEqual(0, errors.Count());
		}
		
		[Test]
		public void AllParagraphProperties()
		{
			using MemoryStream mem = new MemoryStream();
			WordDocument doc = new WordDocument(mem);

			doc.Process(new HtmlParser("<p style='text-align:center;margin:5px;background-color:#ccc;border:1px solid #000'>test</p>"));

			Assert.IsNotNull(doc.Document.Body);
			Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

			Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
			Assert.IsNotNull(paragraph);
			Assert.AreEqual(2, paragraph.ChildElements.Count);

			ParagraphProperties properties = paragraph.ChildElements[0] as ParagraphProperties;
			Assert.IsNotNull(properties);

			ParagraphBorders paragraphBorders = properties.ChildElements[0] as ParagraphBorders;
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

			Assert.IsNotNull(paragraph.ParagraphProperties.Shading);
			Assert.AreEqual("cccccc", paragraph.ParagraphProperties.Shading.Fill.Value);
			Assert.AreEqual(Word.ShadingPatternValues.Clear, paragraph.ParagraphProperties.Shading.Val.Value);

			SpacingBetweenLines spacing = properties.ChildElements[2] as SpacingBetweenLines;
			Assert.IsNotNull(spacing);
			Assert.AreEqual("100", spacing.Before.Value);

			Indentation ind = properties.ChildElements[3] as Indentation;
			Assert.IsNotNull(ind);
			Assert.AreEqual("100", ind.Left.Value);
			Assert.IsNotNull(ind.Right);
			Assert.AreEqual("100", ind.Right.Value);

			Justification align = properties.ChildElements[4] as Justification;
			Assert.IsNotNull(align);
			Assert.AreEqual(JustificationValues.Center, align.Val.Value);

			Run run = paragraph.ChildElements[1] as Run;
			Assert.IsNotNull(run);
			Assert.AreEqual(2, run.ChildElements.Count);

			RunProperties runProperties = run.ChildElements[0] as RunProperties;
			Assert.IsNotNull(runProperties);
			Assert.AreEqual("cccccc", runProperties.Shading.Fill.Value);

			Word.Text text = run.ChildElements[1] as Word.Text;
			Assert.IsNotNull(text);
			Assert.AreEqual(0, text.ChildElements.Count);
			Assert.AreEqual("test", text.InnerText);

			OpenXmlValidator validator = new OpenXmlValidator();
			var errors = validator.Validate(doc.WordprocessingDocument);
			errors.PrintValidationErrors();
			Assert.AreEqual(0, errors.Count());
		}
		
		[Test]
		public void TestAllRunProperties()
		{
			using MemoryStream mem = new MemoryStream();
			WordDocument doc = new WordDocument(mem);

			doc.Process(new HtmlParser("<p><span style='font-family:arial;font-weight:bold;text-decoration:underline;font-size:12px;font-style:italic;background-color:#ccc;color:#000'>test</span></p>"));

			Assert.IsNotNull(doc.Document.Body);
			Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

			Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
			Assert.IsNotNull(paragraph);
			Assert.AreEqual(1, paragraph.ChildElements.Count);

			Run run = paragraph.ChildElements[0] as Run;
			Assert.IsNotNull(run);
			Assert.AreEqual(2, run.ChildElements.Count);

			RunProperties properties = run.ChildElements[0] as RunProperties;
			Assert.IsNotNull(properties);

			RunFonts fonts = properties.ChildElements[0] as RunFonts;
			Assert.IsNotNull(fonts);
			Assert.AreEqual("arial", fonts.Ascii.Value);

			Bold bold = properties.ChildElements[1] as Bold;
			Assert.IsNotNull(bold);

			Italic italic = properties.ChildElements[2] as Italic;
			Assert.IsNotNull(italic);

			Word.Color color = properties.ChildElements[3] as Word.Color;
			Assert.IsNotNull(color);
			Assert.AreEqual("000000", color.Val.Value);

			FontSize fontSize = properties.ChildElements[4] as FontSize;
			Assert.IsNotNull(fontSize);
			Assert.AreEqual("24", fontSize.Val.Value);

			Underline underline = properties.ChildElements[5] as Underline;
			Assert.IsNotNull(underline);

			Word.Shading shading = properties.ChildElements[6] as Word.Shading;
			Assert.IsNotNull(shading);
			Assert.AreEqual("cccccc", shading.Fill.Value);
			Assert.AreEqual(Word.ShadingPatternValues.Clear, shading.Val.Value);

			Word.Text text = run.ChildElements[1] as Word.Text;
			Assert.IsNotNull(text);
			Assert.AreEqual(0, text.ChildElements.Count);
			Assert.AreEqual("test", text.InnerText);

			OpenXmlValidator validator = new OpenXmlValidator();
			var errors = validator.Validate(doc.WordprocessingDocument);
			errors.PrintValidationErrors();
			Assert.AreEqual(0, errors.Count());
		}
		
		[Test]
		public void TestRunBackgroundColor()
		{
			using MemoryStream mem = new MemoryStream();
			WordDocument doc = new WordDocument(mem);

			doc.Process(new HtmlParser("<p><span style='background-color:#ccc;color:#000'>one</span><span>two</span></p>"));

			Assert.IsNotNull(doc.Document.Body);
			Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

			Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
			Assert.IsNotNull(para);
			Assert.AreEqual(2, para.ChildElements.Count);

			Run run = para.ChildElements[0] as Run;
			Assert.IsNotNull(run);
			Assert.AreEqual(2, run.ChildElements.Count);

			RunProperties properties = run.ChildElements[0] as RunProperties;
			Assert.IsNotNull(properties);

			Word.Color color = properties.ChildElements[0] as Word.Color;
			Assert.IsNotNull(color);
			Assert.AreEqual("000000", color.Val.Value);

			Word.Shading shading = properties.ChildElements[1] as Word.Shading;
			Assert.IsNotNull(shading);
			Assert.AreEqual("cccccc", shading.Fill.Value);
			Assert.AreEqual(Word.ShadingPatternValues.Clear, shading.Val.Value);

			Word.Text text = run.ChildElements[1] as Word.Text;
			Assert.IsNotNull(text);
			Assert.AreEqual(0, text.ChildElements.Count);
			Assert.AreEqual("one", text.InnerText);

			run = para.ChildElements[1] as Run;
			Assert.IsNotNull(run);
			Assert.AreEqual(1, run.ChildElements.Count);

			text = run.ChildElements[0] as Word.Text;
			Assert.IsNotNull(text);
			Assert.AreEqual(0, text.ChildElements.Count);
			Assert.AreEqual("two", text.InnerText);

			OpenXmlValidator validator = new OpenXmlValidator();
			var errors = validator.Validate(doc.WordprocessingDocument);
			errors.PrintValidationErrors();
			Assert.AreEqual(0, errors.Count());
		}

        [Test]
        public void TestAlignJustify()
        {
			using MemoryStream mem = new MemoryStream();
			WordDocument doc = new WordDocument(mem);

			doc.Process(new HtmlParser("<p style='text-align:justify;'>test</p>"));

			Assert.IsNotNull(doc.Document.Body);
			Assert.AreEqual(1, doc.Document.Body.ChildElements.Count);

			Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
			Assert.IsNotNull(paragraph);
			Assert.AreEqual(2, paragraph.ChildElements.Count);

			ParagraphProperties properties = paragraph.ChildElements[0] as ParagraphProperties;
			Assert.IsNotNull(properties);

			Justification align = properties.ChildElements[0] as Justification;
			Assert.IsNotNull(align);
			Assert.AreEqual(JustificationValues.Both, align.Val.Value);

			Run run = paragraph.ChildElements[1] as Run;
			Assert.IsNotNull(run);
			Assert.AreEqual(1, run.ChildElements.Count);

			Word.Text text = run.ChildElements[0] as Word.Text;
			Assert.IsNotNull(text);
			Assert.AreEqual(0, text.ChildElements.Count);
			Assert.AreEqual("test", text.InnerText);

			OpenXmlValidator validator = new OpenXmlValidator();
			var errors = validator.Validate(doc.WordprocessingDocument);
			errors.PrintValidationErrors();
			Assert.AreEqual(0, errors.Count());
		}
	}
}
