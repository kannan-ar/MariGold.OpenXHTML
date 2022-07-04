namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class TestP
    {
        [Fact]
        public void SinglePBackGround()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p style='background-color:#000'>test</p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.Equal(2, paragraph.ChildElements.Count);
            Assert.NotNull(paragraph.ParagraphProperties);
            Assert.NotNull(paragraph.ParagraphProperties.Shading);
            Assert.Equal("000000", paragraph.ParagraphProperties.Shading.Fill.Value);
            Assert.Equal(Word.ShadingPatternValues.Clear, paragraph.ParagraphProperties.Shading.Val.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties runProperties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(runProperties);
            Assert.Equal("000000", runProperties.Shading.Fill.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SinglePRedColor()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p style='color:red'>test</p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            Word.Color color = run.RunProperties.ChildElements[0] as Word.Color;
            Assert.Equal("FF0000", color.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SinglePAllBorder()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p style='border:1px solid #000'>test</p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties paragraphProperties = paragraph.ChildElements[0] as ParagraphProperties;
            ParagraphBorders paragraphBorders = paragraphProperties.ChildElements[0] as ParagraphBorders;
            Assert.NotNull(paragraphBorders);
            Assert.Equal(4, paragraphBorders.ChildElements.Count);

            TopBorder topBorder = paragraphBorders.ChildElements[0] as TopBorder;
            Assert.NotNull(topBorder);
            Assert.Equal(BorderValues.Single, topBorder.Val.Value);
            Assert.Equal("000000", topBorder.Color.Value);
            TestUtility.Equal(1, topBorder.Size.Value);

            LeftBorder leftBorder = paragraphBorders.ChildElements[1] as LeftBorder;
            Assert.NotNull(leftBorder);
            Assert.Equal(BorderValues.Single, leftBorder.Val.Value);
            Assert.Equal("000000", leftBorder.Color.Value);
            TestUtility.Equal(1, leftBorder.Size.Value);

            BottomBorder bottomBorder = paragraphBorders.ChildElements[2] as BottomBorder;
            Assert.NotNull(bottomBorder);
            Assert.Equal(BorderValues.Single, bottomBorder.Val.Value);
            Assert.Equal("000000", bottomBorder.Color.Value);
            TestUtility.Equal(1, bottomBorder.Size.Value);

            RightBorder rightBorder = paragraphBorders.ChildElements[3] as RightBorder;
            Assert.NotNull(rightBorder);
            Assert.Equal(BorderValues.Single, rightBorder.Val.Value);
            Assert.Equal("000000", rightBorder.Color.Value);
            TestUtility.Equal(1, rightBorder.Size.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SinglePAllBorderWithBorderBottom()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p style='border:1px solid #000;border-bottom:red solid 2px'>test</p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties paragraphProperties = paragraph.ChildElements[0] as ParagraphProperties;
            ParagraphBorders paragraphBorders = paragraphProperties.ChildElements[0] as ParagraphBorders;
            Assert.NotNull(paragraphBorders);
            Assert.Equal(4, paragraphBorders.ChildElements.Count);

            TopBorder topBorder = paragraphBorders.ChildElements[0] as TopBorder;
            Assert.NotNull(topBorder);
            Assert.Equal(BorderValues.Single, topBorder.Val.Value);
            Assert.Equal("000000", topBorder.Color.Value);
            TestUtility.Equal(1, topBorder.Size.Value);

            LeftBorder leftBorder = paragraphBorders.ChildElements[1] as LeftBorder;
            Assert.NotNull(leftBorder);
            Assert.Equal(BorderValues.Single, leftBorder.Val.Value);
            Assert.Equal("000000", leftBorder.Color.Value);
            TestUtility.Equal(1, leftBorder.Size.Value);

            BottomBorder bottomBorder = paragraphBorders.ChildElements[2] as BottomBorder;
            Assert.NotNull(bottomBorder);
            Assert.Equal(BorderValues.Single, bottomBorder.Val.Value);
            Assert.Equal("FF0000", bottomBorder.Color.Value);
            TestUtility.Equal(2, bottomBorder.Size.Value);

            RightBorder rightBorder = paragraphBorders.ChildElements[3] as RightBorder;
            Assert.NotNull(rightBorder);
            Assert.Equal(BorderValues.Single, rightBorder.Val.Value);
            Assert.Equal("000000", rightBorder.Color.Value);
            TestUtility.Equal(1, rightBorder.Size.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SinglePAllBorderWithIndependentBorders()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p style='border:1px solid #000;border-bottom:red solid 2px;border-top:blue 3px solid;border-left:#F0F000 4px solid;border-right:solid #CCC888 5px'>test</p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties paragraphProperties = paragraph.ChildElements[0] as ParagraphProperties;
            ParagraphBorders paragraphBorders = paragraphProperties.ChildElements[0] as ParagraphBorders;
            Assert.NotNull(paragraphBorders);
            Assert.Equal(4, paragraphBorders.ChildElements.Count);

            TopBorder topBorder = paragraphBorders.ChildElements[0] as TopBorder;
            Assert.NotNull(topBorder);
            Assert.Equal(BorderValues.Single, topBorder.Val.Value);
            Assert.Equal("0000FF", topBorder.Color.Value);
            TestUtility.Equal(3, topBorder.Size.Value);

            LeftBorder leftBorder = paragraphBorders.ChildElements[1] as LeftBorder;
            Assert.NotNull(leftBorder);
            Assert.Equal(BorderValues.Single, leftBorder.Val.Value);
            Assert.Equal("f0f000", leftBorder.Color.Value);
            TestUtility.Equal(4, leftBorder.Size.Value);

            BottomBorder bottomBorder = paragraphBorders.ChildElements[2] as BottomBorder;
            Assert.NotNull(bottomBorder);
            Assert.Equal(BorderValues.Single, bottomBorder.Val.Value);
            Assert.Equal("FF0000", bottomBorder.Color.Value);
            TestUtility.Equal(2, bottomBorder.Size.Value);

            RightBorder rightBorder = paragraphBorders.ChildElements[3] as RightBorder;
            Assert.NotNull(rightBorder);
            Assert.Equal(BorderValues.Single, rightBorder.Val.Value);
            Assert.Equal("ccc888", rightBorder.Color.Value);
            TestUtility.Equal(5, rightBorder.Size.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);
            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void SinglePFontSize()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p style='font-size:10px'>test</p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            FontSize fontSize = run.RunProperties.ChildElements[0] as FontSize;
            Assert.Equal("20", fontSize.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void AllParagraphProperties()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p style='text-align:center;margin:5px;background-color:#ccc;border:1px solid #000'>test</p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties properties = paragraph.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);

            ParagraphBorders paragraphBorders = properties.ChildElements[0] as ParagraphBorders;
            Assert.NotNull(paragraphBorders);
            Assert.Equal(4, paragraphBorders.ChildElements.Count);

            TopBorder topBorder = paragraphBorders.ChildElements[0] as TopBorder;
            Assert.NotNull(topBorder);
            Assert.Equal(BorderValues.Single, topBorder.Val.Value);
            Assert.Equal("000000", topBorder.Color.Value);
            TestUtility.Equal(1, topBorder.Size.Value);

            LeftBorder leftBorder = paragraphBorders.ChildElements[1] as LeftBorder;
            Assert.NotNull(leftBorder);
            Assert.Equal(BorderValues.Single, leftBorder.Val.Value);
            Assert.Equal("000000", leftBorder.Color.Value);
            TestUtility.Equal(1, leftBorder.Size.Value);

            BottomBorder bottomBorder = paragraphBorders.ChildElements[2] as BottomBorder;
            Assert.NotNull(bottomBorder);
            Assert.Equal(BorderValues.Single, bottomBorder.Val.Value);
            Assert.Equal("000000", bottomBorder.Color.Value);
            TestUtility.Equal(1, bottomBorder.Size.Value);

            RightBorder rightBorder = paragraphBorders.ChildElements[3] as RightBorder;
            Assert.NotNull(rightBorder);
            Assert.Equal(BorderValues.Single, rightBorder.Val.Value);
            Assert.Equal("000000", rightBorder.Color.Value);
            TestUtility.Equal(1, rightBorder.Size.Value);

            Assert.NotNull(paragraph.ParagraphProperties.Shading);
            Assert.Equal("cccccc", paragraph.ParagraphProperties.Shading.Fill.Value);
            Assert.Equal(Word.ShadingPatternValues.Clear, paragraph.ParagraphProperties.Shading.Val.Value);

            SpacingBetweenLines spacing = properties.ChildElements[2] as SpacingBetweenLines;
            Assert.NotNull(spacing);
            Assert.Equal("100", spacing.Before.Value);

            Indentation ind = properties.ChildElements[3] as Indentation;
            Assert.NotNull(ind);
            Assert.Equal("100", ind.Left.Value);
            Assert.NotNull(ind.Right);
            Assert.Equal("100", ind.Right.Value);

            Justification align = properties.ChildElements[4] as Justification;
            Assert.NotNull(align);
            Assert.Equal(JustificationValues.Center, align.Val.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties runProperties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(runProperties);
            Assert.Equal("cccccc", runProperties.Shading.Fill.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void TestAllRunProperties()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p><span style='font-family:arial;font-weight:bold;text-decoration:underline;font-size:12px;font-style:italic;background-color:#ccc;color:#000'>test</span></p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(1, paragraph.ChildElements.Count);

            Run run = paragraph.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);

            RunFonts fonts = properties.ChildElements[0] as RunFonts;
            Assert.NotNull(fonts);
            Assert.Equal("arial", fonts.Ascii.Value);

            Bold bold = properties.ChildElements[1] as Bold;
            Assert.NotNull(bold);

            Italic italic = properties.ChildElements[2] as Italic;
            Assert.NotNull(italic);

            Word.Color color = properties.ChildElements[3] as Word.Color;
            Assert.NotNull(color);
            Assert.Equal("000000", color.Val.Value);

            FontSize fontSize = properties.ChildElements[4] as FontSize;
            Assert.NotNull(fontSize);
            Assert.Equal("24", fontSize.Val.Value);

            Underline underline = properties.ChildElements[5] as Underline;
            Assert.NotNull(underline);

            Word.Shading shading = properties.ChildElements[6] as Word.Shading;
            Assert.NotNull(shading);
            Assert.Equal("cccccc", shading.Fill.Value);
            Assert.Equal(Word.ShadingPatternValues.Clear, shading.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void TestRunBackgroundColor()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p><span style='background-color:#ccc;color:#000'>one</span><span>two</span></p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph para = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(2, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties properties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(properties);

            Word.Color color = properties.ChildElements[0] as Word.Color;
            Assert.NotNull(color);
            Assert.Equal("000000", color.Val.Value);

            Word.Shading shading = properties.ChildElements[1] as Word.Shading;
            Assert.NotNull(shading);
            Assert.Equal("cccccc", shading.Fill.Value);
            Assert.Equal(Word.ShadingPatternValues.Clear, shading.Val.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("one", text.InnerText);

            run = para.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("two", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void TestAlignJustify()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<p style='text-align:justify;'>test</p>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Paragraph paragraph = doc.Document.Body.ChildElements[0] as Paragraph;
            Assert.NotNull(paragraph);
            Assert.Equal(2, paragraph.ChildElements.Count);

            ParagraphProperties properties = paragraph.ChildElements[0] as ParagraphProperties;
            Assert.NotNull(properties);

            Justification align = properties.ChildElements[0] as Justification;
            Assert.NotNull(align);
            Assert.Equal(JustificationValues.Both, align.Val.Value);

            Run run = paragraph.ChildElements[1] as Run;
            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }
    }
}
