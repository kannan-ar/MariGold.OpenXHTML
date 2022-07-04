namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
    using OpenXHTML;
    using System.IO;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    public class Table
    {
        [Fact]
        public void TableBorder()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table border='1'><tr><td>test</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableBorders tableBorders = tableProperties.ChildElements[1] as TableBorders;
            Assert.NotNull(tableBorders);
            Assert.Equal(4, tableBorders.ChildElements.Count);

            TopBorder topBorder = tableBorders.ChildElements[0] as TopBorder;
            Assert.NotNull(topBorder);
            TestUtility.TestBorder<TopBorder>(topBorder, BorderValues.Single, "auto", 4U);

            LeftBorder leftBorder = tableBorders.ChildElements[1] as LeftBorder;
            Assert.NotNull(leftBorder);
            TestUtility.TestBorder<LeftBorder>(leftBorder, BorderValues.Single, "auto", 4U);

            BottomBorder bottomBorder = tableBorders.ChildElements[2] as BottomBorder;
            Assert.NotNull(bottomBorder);
            TestUtility.TestBorder<BottomBorder>(bottomBorder, BorderValues.Single, "auto", 4U);

            RightBorder rightBorder = tableBorders.ChildElements[3] as RightBorder;
            Assert.NotNull(rightBorder);
            TestUtility.TestBorder<RightBorder>(rightBorder, BorderValues.Single, "auto", 4U);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(2, cell.ChildElements.Count);

            TableCellProperties cellProperties = cell.ChildElements[0] as TableCellProperties;
            Assert.NotNull(cellProperties);
            Assert.Equal(1, cellProperties.ChildElements.Count);

            TableCellBorders cellBorders = cellProperties.ChildElements[0] as TableCellBorders;
            Assert.NotNull(cellBorders);
            Assert.Equal(4, cellBorders.ChildElements.Count);

            topBorder = cellBorders.ChildElements[0] as TopBorder;
            Assert.NotNull(topBorder);
            TestUtility.TestBorder<TopBorder>(topBorder, BorderValues.Single, "auto", 4U);

            leftBorder = cellBorders.ChildElements[1] as LeftBorder;
            Assert.NotNull(leftBorder);
            TestUtility.TestBorder<LeftBorder>(leftBorder, BorderValues.Single, "auto", 4U);

            bottomBorder = cellBorders.ChildElements[2] as BottomBorder;
            Assert.NotNull(bottomBorder);
            TestUtility.TestBorder<BottomBorder>(bottomBorder, BorderValues.Single, "auto", 4U);

            rightBorder = cellBorders.ChildElements[3] as RightBorder;
            Assert.NotNull(rightBorder);
            TestUtility.TestBorder<RightBorder>(rightBorder, BorderValues.Single, "auto", 4U);

            Paragraph para = cell.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

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
        public void TableBorderStyle()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table style='border:1px solid #000'><tr><td>test</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableBorders tableBorders = tableProperties.ChildElements[1] as TableBorders;
            Assert.NotNull(tableBorders);
            Assert.Equal(4, tableBorders.ChildElements.Count);

            TopBorder topBorder = tableBorders.ChildElements[0] as TopBorder;
            Assert.NotNull(topBorder);
            TestUtility.TestBorder<TopBorder>(topBorder, BorderValues.Single, "000000", 1U);

            LeftBorder leftBorder = tableBorders.ChildElements[1] as LeftBorder;
            Assert.NotNull(leftBorder);
            TestUtility.TestBorder<LeftBorder>(leftBorder, BorderValues.Single, "000000", 1U);

            BottomBorder bottomBorder = tableBorders.ChildElements[2] as BottomBorder;
            Assert.NotNull(bottomBorder);
            TestUtility.TestBorder<BottomBorder>(bottomBorder, BorderValues.Single, "000000", 1U);

            RightBorder rightBorder = tableBorders.ChildElements[3] as RightBorder;
            Assert.NotNull(rightBorder);
            TestUtility.TestBorder<RightBorder>(rightBorder, BorderValues.Single, "000000", 1U);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

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
        public void TableCellSpacing()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table cellspacing='5'><tr><td>test</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(2, row.ChildElements.Count);

            TableRowProperties rowProperties = row.ChildElements[0] as TableRowProperties;
            Assert.NotNull(rowProperties);
            Assert.Equal(1, rowProperties.ChildElements.Count);

            TableCellSpacing cellSpacing = rowProperties.ChildElements[0] as TableCellSpacing;
            Assert.NotNull(cellSpacing);
            Assert.Equal("100", cellSpacing.Width.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, cellSpacing.Type.Value);

            TableCell cell = row.ChildElements[1] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

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
        public void TableCellPadding()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table cellpadding='5'><tr><td>test</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);
            Assert.NotNull(tableProperties.TableCellMarginDefault);

            Assert.Equal(100, tableProperties.TableCellMarginDefault.TableCellLeftMargin.Width.Value);
            Assert.Equal(TableWidthValues.Dxa, tableProperties.TableCellMarginDefault.TableCellLeftMargin.Type.Value);

            Assert.Equal("100", tableProperties.TableCellMarginDefault.TopMargin.Width.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, tableProperties.TableCellMarginDefault.TopMargin.Type.Value);

            Assert.Equal(100, tableProperties.TableCellMarginDefault.TableCellRightMargin.Width.Value);
            Assert.Equal(TableWidthValues.Dxa, tableProperties.TableCellMarginDefault.TableCellRightMargin.Type.Value);

            Assert.Equal("100", tableProperties.TableCellMarginDefault.BottomMargin.Width.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, tableProperties.TableCellMarginDefault.BottomMargin.Type.Value);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

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
        public void TableThCellStyles()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><th>Id</th></tr><tr><td>1</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(4, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableGrid tableGrid = table.ChildElements[1] as TableGrid;
            Assert.NotNull(tableGrid);
            Assert.Equal(1, tableGrid.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            Bold bold = run.RunProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("Id", text.InnerText);

            row = table.ChildElements[3] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            cell = row.ChildElements[0] as TableCell;
            cell.TestTableCell(1, "1");

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TableThColSpan()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><td>Id</td><td>Name</td></tr><tr><td colspan='2'>1</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(4, table.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(2, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);
            cell.TestTableCell(1, "Id");

            cell = row.ChildElements[1] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);
            cell.TestTableCell(1, "Name");

            row = table.ChildElements[3] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);


            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TableThColSpanAdv()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table border='1'><tr><td>1</td><td>2</td><td>3</td><td>4</td><td>5</td></tr><tr><td colspan='2'>one</td><td>three</td><td colspan='2'>five</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(4, table.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(5, row.ChildElements.Count);

            row = table.ChildElements[3] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(3, row.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TableAttributeWidth()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table width='50%'><tr><td>test</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);
            Assert.Equal(2, tableProperties.ChildElements.Count);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableWidth tableWidth = tableProperties.ChildElements[1] as TableWidth;
            Assert.NotNull(tableWidth);
            Assert.Equal("2500", tableWidth.Width.Value);
            Assert.Equal(TableWidthUnitValues.Pct, tableWidth.Type.Value);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

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
        public void TableStyleWidth()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table style='width:50%'><tr><td>test</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);
            Assert.Equal(2, tableProperties.ChildElements.Count);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableWidth tableWidth = tableProperties.ChildElements[1] as TableWidth;
            Assert.NotNull(tableWidth);
            Assert.Equal("2500", tableWidth.Width.Value);
            Assert.Equal(TableWidthUnitValues.Pct, tableWidth.Type.Value);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

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
        public void TableAttributeStyleWidth()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table width='50%' style='width:150px'><tr><td>test</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);
            Assert.Equal(2, tableProperties.ChildElements.Count);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableWidth tableWidth = tableProperties.ChildElements[1] as TableWidth;
            Assert.NotNull(tableWidth);
            Assert.Equal("3000", tableWidth.Width.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, tableWidth.Type.Value);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

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
        public void TableCellWidth()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table style='width:500px'><tr><td style='width:250px'>1</td><td style='width:250px'>2</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);
            Assert.Equal(2, tableProperties.ChildElements.Count);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableWidth tableWidth = tableProperties.ChildElements[1] as TableWidth;
            Assert.NotNull(tableWidth);
            Assert.Equal("10000", tableWidth.Width.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, tableWidth.Type.Value);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(2, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(2, cell.ChildElements.Count);

            TableCellProperties cellProperties = cell.ChildElements[0] as TableCellProperties;
            Assert.NotNull(cellProperties);
            TableCellWidth cellWidth = cellProperties.ChildElements[0] as TableCellWidth;
            Assert.NotNull(cellWidth);
            Assert.Equal("5000", cellWidth.Width.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, cellWidth.Type.Value);

            Paragraph para = cell.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("1", text.InnerText);

            cell = row.ChildElements[1] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(2, cell.ChildElements.Count);

            cellProperties = cell.ChildElements[0] as TableCellProperties;
            Assert.NotNull(cellProperties);
            cellWidth = cellProperties.ChildElements[0] as TableCellWidth;
            Assert.NotNull(cellWidth);
            Assert.Equal("5000", cellWidth.Width.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, cellWidth.Type.Value);

            para = cell.ChildElements[1] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("2", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TableAllProperties()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table border='1' style='width:500px' cellpadding='2'><tr><td>1</td><td>2</td></tr><tr><td colspan='2'>1</td></tr></table>"));

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            errors.PrintValidationErrors();
            Assert.Empty(errors);
        }

        [Fact]
        public void EmptyCell()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><td></td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);
            Assert.Equal(1, tableProperties.ChildElements.Count);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableGrid tableGrid = table.ChildElements[1] as TableGrid;
            Assert.NotNull(tableGrid);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(0, para.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void RowSpan()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><td rowspan='2'></td><td></td></tr><tr><td></td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(4, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);
            Assert.Equal(1, tableProperties.ChildElements.Count);

            TableGrid tableGrid = table.ChildElements[1] as TableGrid;
            Assert.NotNull(tableGrid);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(2, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(2, cell.ChildElements.Count);

            TableCellProperties cellProperties = cell.ChildElements[0] as TableCellProperties;
            Assert.NotNull(cellProperties);
            Assert.Equal(1, cellProperties.ChildElements.Count);
            VerticalMerge verticalMerge = cellProperties.ChildElements[0] as VerticalMerge;
            Assert.NotNull(verticalMerge);
            Assert.Equal(MergedCellValues.Restart, verticalMerge.Val.Value);

            Paragraph para = cell.ChildElements[1] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(0, para.ChildElements.Count);

            cell = row.ChildElements[1] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            para = cell.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(0, para.ChildElements.Count);

            row = table.ChildElements[3] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(2, row.ChildElements.Count);

            cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(2, cell.ChildElements.Count);

            cellProperties = cell.ChildElements[0] as TableCellProperties;
            Assert.NotNull(cellProperties);
            Assert.Equal(1, cellProperties.ChildElements.Count);
            verticalMerge = cellProperties.ChildElements[0] as VerticalMerge;
            Assert.NotNull(verticalMerge);
            Assert.Equal(false, verticalMerge.HasChildren);

            para = cell.ChildElements[1] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(0, para.ChildElements.Count);

            cell = row.ChildElements[1] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            para = cell.ChildElements[0] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(0, para.ChildElements.Count);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TableCellStyleInheritance()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><td style='border:1px solid #000;background-color:red'>test</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(2, cell.ChildElements.Count);

            TableCellProperties cellProperties = cell.ChildElements[0] as TableCellProperties;

            Assert.NotNull(cellProperties);
            Assert.Equal(2, cellProperties.ChildElements.Count);

            TableCellBorders borders = cellProperties.ChildElements[0] as TableCellBorders;
            Assert.NotNull(borders);
            Assert.Equal(4, borders.ChildElements.Count);

            TopBorder topBorder = borders.ChildElements[0] as TopBorder;
            Assert.NotNull(topBorder);
            TestUtility.TestBorder<TopBorder>(topBorder, BorderValues.Single, "000000", 1U);

            LeftBorder leftBorder = borders.ChildElements[1] as LeftBorder;
            Assert.NotNull(leftBorder);
            TestUtility.TestBorder<LeftBorder>(leftBorder, BorderValues.Single, "000000", 1U);

            BottomBorder bottomBorder = borders.ChildElements[2] as BottomBorder;
            Assert.NotNull(bottomBorder);
            TestUtility.TestBorder<BottomBorder>(bottomBorder, BorderValues.Single, "000000", 1U);

            RightBorder rightBorder = borders.ChildElements[3] as RightBorder;
            Assert.NotNull(rightBorder);
            TestUtility.TestBorder<RightBorder>(rightBorder, BorderValues.Single, "000000", 1U);

            Word.Shading backgroundColor = cellProperties.ChildElements[1] as Word.Shading;
            Assert.NotNull(backgroundColor);
            Assert.Equal("FF0000", backgroundColor.Fill.Value);

            Paragraph para = cell.ChildElements[1] as Paragraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);

            RunProperties runProperties = run.ChildElements[0] as RunProperties;
            Assert.NotNull(runProperties);
            Word.Shading shading = runProperties.ChildElements[0] as Word.Shading;
            Assert.Equal("FF0000", shading.Fill.Value);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal("test", text.InnerText);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TableInsideAnotherTable()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><td><table><tr><td>test</td></tr></table></td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(3, table.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(2, cell.ChildElements.Count);

            table = cell.ChildElements[0] as Word.Table;
            Assert.NotNull(table);
            Paragraph para = cell.ChildElements[1] as Paragraph;
            Assert.NotNull(para);

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TableWithTBody()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><th>Id</th></tr><tbody><tr><td>1</td></tr></tbody></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(4, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableGrid tableGrid = table.ChildElements[1] as TableGrid;
            Assert.NotNull(tableGrid);
            Assert.Equal(1, tableGrid.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            Bold bold = run.RunProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("Id", text.InnerText);

            row = table.ChildElements[3] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            cell = row.ChildElements[0] as TableCell;
            cell.TestTableCell(1, "1");

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TableWithTHead()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><thead><tr><th>Id</th></tr></thead><tr><td>1</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(4, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableGrid tableGrid = table.ChildElements[1] as TableGrid;
            Assert.NotNull(tableGrid);
            Assert.Equal(1, tableGrid.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            Bold bold = run.RunProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("Id", text.InnerText);

            row = table.ChildElements[3] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            cell = row.ChildElements[0] as TableCell;
            cell.TestTableCell(1, "1");

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TableWithTFoot()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><thead><tr><th>Id</th></tr></thead><tfoot><tr><td>1</td></tr></tfoot></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(4, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableGrid tableGrid = table.ChildElements[1] as TableGrid;
            Assert.NotNull(tableGrid);
            Assert.Equal(1, tableGrid.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            Bold bold = run.RunProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("Id", text.InnerText);

            row = table.ChildElements[3] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            cell = row.ChildElements[0] as TableCell;
            cell.TestTableCell(1, "1");

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void ColSpanTableGrid()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><tr><td colspan='2'>head</td></tr><tr><td>1</td><td>2</td></tr></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(4, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);
            Assert.Equal(1, tableProperties.ChildElements.Count);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableGrid tableGrid = table.ChildElements[1] as TableGrid;
            Assert.NotNull(tableGrid);
            Assert.Equal(2, tableGrid.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;

            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;

            Assert.NotNull(cell);
            Assert.Equal(2, cell.ChildElements.Count);

            TableCellProperties cellProperties = cell.ChildElements[0] as TableCellProperties;

            Assert.NotNull(cellProperties);
            Assert.Equal(1, cellProperties.ChildElements.Count);

            GridSpan gridSpan = cellProperties.ChildElements[0] as GridSpan;
            Assert.NotNull(gridSpan);
            Assert.Equal(2, gridSpan.Val.Value);

            Word.Paragraph para = cell.ChildElements[1] as Word.Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Word.Run run = para.ChildElements[0] as Word.Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("head", text.InnerText);

            row = table.ChildElements[3] as TableRow;

            cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            cell.TestTableCell(1, "1");

            cell = row.ChildElements[1] as TableCell;
            Assert.NotNull(cell);
            cell.TestTableCell(1, "2");

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }

        [Fact]
        public void TableWithTHeadTBodyTFoot()
        {
            using MemoryStream mem = new MemoryStream();
            WordDocument doc = new WordDocument(mem);

            doc.Process(new HtmlParser("<table><thead><tr><th>Id</th></tr></thead><tbody><tr><td>1</td></tr></tbody><tfoot><tr><td>2</td></tr></tfoot></table>"));

            Assert.NotNull(doc.Document.Body);
            Assert.Equal(1, doc.Document.Body.ChildElements.Count);

            Word.Table table = doc.Document.Body.ChildElements[0] as Word.Table;

            Assert.NotNull(table);
            Assert.Equal(5, table.ChildElements.Count);

            TableProperties tableProperties = table.ChildElements[0] as TableProperties;
            Assert.NotNull(tableProperties);

            TableStyle tableStyle = tableProperties.ChildElements[0] as TableStyle;
            Assert.NotNull(tableStyle);
            Assert.Equal("TableGrid", tableStyle.Val.Value);

            TableGrid tableGrid = table.ChildElements[1] as TableGrid;
            Assert.NotNull(tableGrid);
            Assert.Equal(1, tableGrid.ChildElements.Count);

            TableRow row = table.ChildElements[2] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            TableCell cell = row.ChildElements[0] as TableCell;
            Assert.NotNull(cell);
            Assert.Equal(1, cell.ChildElements.Count);

            Paragraph para = cell.ChildElements[0] as Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Run run = para.ChildElements[0] as Run;

            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            Assert.NotNull(run.RunProperties);
            Bold bold = run.RunProperties.ChildElements[0] as Bold;
            Assert.NotNull(bold);

            Word.Text text = run.ChildElements[1] as Word.Text;
            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal("Id", text.InnerText);

            row = table.ChildElements[3] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            cell = row.ChildElements[0] as TableCell;
            cell.TestTableCell(1, "1");

            row = table.ChildElements[4] as TableRow;
            Assert.NotNull(row);
            Assert.Equal(1, row.ChildElements.Count);

            cell = row.ChildElements[0] as TableCell;
            cell.TestTableCell(1, "2");

            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.WordprocessingDocument);
            Assert.Empty(errors);
        }
    }
}
