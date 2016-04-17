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
		
		internal bool HasRowSpan{ get; set; }
		
		internal void Process(TableCell cell, DocxTableProperties docxProperties, IHtmlNode node)
		{
			DocxNode docxNode = new DocxNode(node);
			TableCellProperties cellProperties = new TableCellProperties();
			
			ProcessColSpan(docxNode, cellProperties);
			ProcessWidth(docxNode, cellProperties);
			
			if (HasRowSpan)
			{
				cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
			}
			
			//Processing border should be after colspan
			ProcessBorders(docxNode, docxProperties, cellProperties);
			
			string backgroundColor = docxNode.ExtractStyleValue(DocxColor.backGroundColor);
			
			if (!string.IsNullOrEmpty(backgroundColor))
			{
				DocxColor.ApplyBackGroundColor(backgroundColor, cellProperties);
			}
			
			ProcessVerticalAlignment(docxNode, cellProperties);
			
			if (cellProperties.HasChildren)
			{
				cell.Append(cellProperties);
			}
		}
		
		internal static IHtmlNode GetHtmlNodeForTableCellContent(IHtmlNode node)
		{
			IHtmlNode clone = node.Clone();
			
			clone.Styles.Remove(DocxBorder.borderName);
			clone.Styles.Remove(DocxBorder.leftBorderName);
			clone.Styles.Remove(DocxBorder.rightBorderName);
			clone.Styles.Remove(DocxBorder.topBorderName);
			clone.Styles.Remove(DocxBorder.bottomBorderName);
			
			clone.Styles.Remove(DocxMargin.margin);
			clone.Styles.Remove(DocxMargin.marginLeft);
			clone.Styles.Remove(DocxMargin.marginRight);
			clone.Styles.Remove(DocxMargin.marginTop);
			clone.Styles.Remove(DocxMargin.marginBottom);
			
			return clone;
		}
	}
}
