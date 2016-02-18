namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal sealed class DocxTableStyle
	{
		private void ProcessTableBorder(DocxNode docxNode, DocxTableProperties docxProperties, TableProperties tableProperties)
		{
			string borderStyle = docxNode.ExtractAttributeValue(DocxBorder.borderName);
			
			if (borderStyle == "1")
			{
				TableBorders tableBorders = new TableBorders();
				DocxBorder.ApplyDefaultBorders(tableBorders);
				tableProperties.Append(tableBorders);
			}
			else
			{
				borderStyle = docxNode.ExtractStyleValue(DocxBorder.borderName);
				string leftBorder = docxNode.ExtractStyleValue(DocxBorder.leftBorderName);
				string topBorder = docxNode.ExtractStyleValue(DocxBorder.topBorderName);
				string rightBorder = docxNode.ExtractStyleValue(DocxBorder.rightBorderName);
				string bottomBorder = docxNode.ExtractStyleValue(DocxBorder.bottomBorderName);
				
				TableBorders tableBorders = new TableBorders();
					
				DocxBorder.ApplyBorders(tableBorders, borderStyle, leftBorder, topBorder, 
					rightBorder, bottomBorder, docxProperties.HasDefaultBorder);
				
				if (tableBorders.HasChildren)
				{
					tableProperties.Append(tableBorders);
				}
			}
		}
		
		private void ProcessTableCellMargin(DocxTableProperties docxProperties, TableProperties tableProperties)
		{
			if (docxProperties.CellPadding != null)
			{
				TableCellMargin cellMargin = new TableCellMargin();
				StringValue width = DocxUnits.GetDxaFromPixel(docxProperties.CellPadding.Value);
				
				cellMargin.LeftMargin = new LeftMargin() {
					Width = width,
					Type = TableWidthUnitValues.Dxa
				};
				
				cellMargin.TopMargin = new TopMargin() {
					Width = width,
					Type = TableWidthUnitValues.Dxa
				};
				
				cellMargin.RightMargin = new RightMargin() {
					Width = width,
					Type = TableWidthUnitValues.Dxa
				};
				
				cellMargin.BottomMargin = new BottomMargin() {
					Width = width,
					Type = TableWidthUnitValues.Dxa
				};
				
				tableProperties.Append(cellMargin);
			}
		}
		
		internal void Process(TableProperties tableProperties, DocxTableProperties docxProperties, IHtmlNode node)
		{
			DocxNode docxNode = new DocxNode(node);
			
			ProcessTableBorder(docxNode, docxProperties, tableProperties);
			
			//	ProcessTableCellMargin(docxProperties, tableProperties);
		}
		
		internal void Process(TableRow row, DocxTableProperties docxProperties, IHtmlNode node)
		{
			TableRowProperties trProperties = new TableRowProperties();
			
			if (docxProperties.CellSpacing != null)
			{
				trProperties.Append(new TableCellSpacing() {
					Width = DocxUnits.GetDxaFromPixel(docxProperties.CellSpacing.Value),
					Type = TableWidthUnitValues.Dxa
				});
			}
			
			if (trProperties.ChildElements.Count > 0)
			{
				row.Append(trProperties);
			}
		}
		
		internal void Process(TableCell cell, DocxTableProperties docxProperties, IHtmlNode node)
		{
			DocxNode docxNode = new DocxNode(node);
			
			string borderStyle = docxNode.ExtractStyleValue(DocxBorder.borderName);
			string leftBorder = docxNode.ExtractStyleValue(DocxBorder.leftBorderName);
			string topBorder = docxNode.ExtractStyleValue(DocxBorder.topBorderName);
			string rightBorder = docxNode.ExtractStyleValue(DocxBorder.rightBorderName);
			string bottomBorder = docxNode.ExtractStyleValue(DocxBorder.bottomBorderName);
			
			TableCellProperties cellProperties = new TableCellProperties();
			TableCellBorders cellBorders = new TableCellBorders();
			
			DocxBorder.ApplyBorders(cellBorders, borderStyle, leftBorder, topBorder, 
				rightBorder, bottomBorder, docxProperties.HasDefaultBorder);
			
			if (cellBorders.HasChildren)
			{
				cellProperties.Append(cellBorders);
			}
			
			if (cellProperties.HasChildren)
			{
				cell.Append(cellProperties);
			}
		}
	}
}
