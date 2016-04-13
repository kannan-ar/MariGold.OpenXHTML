namespace MariGold.OpenXHTML
{
	using System;
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
				TableCellMarginDefault cellMargin = new TableCellMarginDefault();
				Int16 width = (Int16)DocxUnits.GetDxaFromPixel(docxProperties.CellPadding.Value);
				
				cellMargin.TableCellLeftMargin = new TableCellLeftMargin() {
					Width = width,
					Type = TableWidthValues.Dxa
				};
				
				cellMargin.TopMargin = new TopMargin() {
					Width = width.ToString(),
					Type = TableWidthUnitValues.Dxa
				};
				
				cellMargin.TableCellRightMargin = new TableCellRightMargin() {
					Width = width,
					Type = TableWidthValues.Dxa
				};
				
				cellMargin.BottomMargin = new BottomMargin() {
					Width = width.ToString(),
					Type = TableWidthUnitValues.Dxa
				};
				
				tableProperties.Append(cellMargin);
			}
		}
		
		private void ProcessWidth(DocxNode docxNode, TableProperties tableProperties)
		{
			string width = docxNode.ExtractAttributeValue(DocxUnits.width);
			string styleWidth = docxNode.ExtractStyleValue(DocxUnits.width);
			
			if (!string.IsNullOrEmpty(styleWidth))
			{
				width = styleWidth;
			}
			
			if (!string.IsNullOrEmpty(width))
			{
				Int32 value;
				TableWidthUnitValues unit;
				
				if (DocxUnits.TableUnitsFromStyle(width, out value, out unit))
				{
					TableWidth tableWidth = new TableWidth() {
						Width = value.ToString(),
						Type = unit
					};
					tableProperties.Append(tableWidth);
				}
			}
		}
		
		internal void Process(TableProperties tableProperties, DocxTableProperties docxProperties, IHtmlNode node)
		{
			DocxNode docxNode = new DocxNode(node);
			ProcessWidth(docxNode, tableProperties);
			
			ProcessTableBorder(docxNode, docxProperties, tableProperties);
			ProcessTableCellMargin(docxProperties, tableProperties);
			
		}
	}
}
