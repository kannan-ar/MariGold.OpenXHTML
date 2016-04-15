namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal sealed class DocxTableCellStyle
	{
		private const string colspan = "colspan";
		
		private void ProcessBorders(DocxNode docxNode, DocxTableProperties docxProperties,
			TableCellProperties cellProperties)
		{
			string borderStyle = docxNode.ExtractStyleValue(DocxBorder.borderName);
			string leftBorder = docxNode.ExtractStyleValue(DocxBorder.leftBorderName);
			string topBorder = docxNode.ExtractStyleValue(DocxBorder.topBorderName);
			string rightBorder = docxNode.ExtractStyleValue(DocxBorder.rightBorderName);
			string bottomBorder = docxNode.ExtractStyleValue(DocxBorder.bottomBorderName);
			
			TableCellBorders cellBorders = new TableCellBorders();
			
			DocxBorder.ApplyBorders(cellBorders, borderStyle, leftBorder, topBorder, 
				rightBorder, bottomBorder, docxProperties.HasDefaultBorder);
			
			if (cellBorders.HasChildren)
			{
				cellProperties.Append(cellBorders);
			}
		}
		
		private void ProcessColSpan(DocxNode docxNode, TableCellProperties cellProperties)
		{
			Int32 value;
			
			if (Int32.TryParse(docxNode.ExtractAttributeValue(colspan), out value))
			{
				if (value > 1)
				{
					cellProperties.Append(new GridSpan() { Val = value });
				}
			}
		}
		
		private void ProcessWidth(DocxNode docxNode, TableCellProperties cellProperties)
		{
			string width = docxNode.ExtractStyleValue(DocxUnits.width);
			
			if (!string.IsNullOrEmpty(width))
			{
				Int32 value;
				TableWidthUnitValues unit;
				
				if (DocxUnits.TableUnitsFromStyle(width, out value, out unit))
				{
					TableCellWidth cellWidth = new TableCellWidth() {
						Width = value.ToString(),
						Type = unit
					};
					
					cellProperties.Append(cellWidth);
				}

			}
		}
		
		private void ProcessVerticalAlignment(DocxNode docxNode, TableCellProperties cellProperties)
		{
			string alignment = docxNode.ExtractStyleValue(DocxAlignment.verticalAlign);
			
			if (!string.IsNullOrEmpty(alignment))
			{
				TableVerticalAlignmentValues value;
				
				if (DocxAlignment.GetCellVerticalAlignment(alignment, out value))
				{
					cellProperties.Append(new TableCellVerticalAlignment(){ Val = value });
				}
			}
		}
		
		private void ProcessVerticalSpan(
			int colIndex, 
			DocxNode docxNode, 
			DocxTableProperties docxProperties, 
			TableCellProperties cellProperties, 
			IHtmlNode node)
		{
			string rowSpan = docxNode.ExtractAttributeValue(DocxTableProperties.rowSpan);
			Int32 rowSpanValue;
			if (Int32.TryParse(rowSpan, out rowSpanValue))
			{
				docxProperties.RowSpanNode[colIndex] = node;
				docxProperties.RowSpanInfo[colIndex] = rowSpanValue - 1;
				cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
			}
		}
		
		internal void Process(int colIndex, TableCell cell, DocxTableProperties docxProperties, IHtmlNode node)
		{
			DocxNode docxNode = new DocxNode(node);
			TableCellProperties cellProperties = new TableCellProperties();
			
			ProcessColSpan(docxNode, cellProperties);
			ProcessWidth(docxNode, cellProperties);
			
			ProcessVerticalSpan(colIndex, docxNode, docxProperties, cellProperties, node);
			//Processing border should be after colspan
			ProcessBorders(docxNode, docxProperties, cellProperties);
			
			ProcessVerticalAlignment(docxNode, cellProperties);
			
			if (cellProperties.HasChildren)
			{
				cell.Append(cellProperties);
			}
		}
	}
}
