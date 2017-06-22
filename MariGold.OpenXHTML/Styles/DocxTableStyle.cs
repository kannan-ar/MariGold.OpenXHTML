namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxTableStyle
	{
        private void ProcessTableBorder(DocxNode node, DocxTableProperties docxProperties, TableProperties tableProperties)
		{
            string borderStyle = node.ExtractAttributeValue(DocxBorder.borderName);
			
			if (borderStyle == "1")
			{
				TableBorders tableBorders = new TableBorders();
				DocxBorder.ApplyDefaultBorders(tableBorders);
				tableProperties.Append(tableBorders);
			}
			else
			{
                borderStyle = node.ExtractStyleValue(DocxBorder.borderName);
                string leftBorder = node.ExtractStyleValue(DocxBorder.leftBorderName);
                string topBorder = node.ExtractStyleValue(DocxBorder.topBorderName);
                string rightBorder = node.ExtractStyleValue(DocxBorder.rightBorderName);
                string bottomBorder = node.ExtractStyleValue(DocxBorder.bottomBorderName);
				
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

        private void ProcessWidth(DocxNode node, TableProperties tableProperties)
		{
            string width = node.ExtractAttributeValue(DocxUnits.width);
            string styleWidth = node.ExtractStyleValue(DocxUnits.width);
			
			if (!string.IsNullOrEmpty(styleWidth))
			{
				width = styleWidth;
			}
			
			if (!string.IsNullOrEmpty(width))
			{
				decimal value;
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
		
		internal void Process(TableProperties tableProperties, DocxTableProperties docxProperties, DocxNode node)
		{
            ProcessWidth(node, tableProperties);

            ProcessTableBorder(node, docxProperties, tableProperties);
			ProcessTableCellMargin(docxProperties, tableProperties);
		}
	}
}
