namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal sealed class DocxTableCellStyle
	{
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
			
			Int32 colSpan;
			
			if (Int32.TryParse(docxNode.ExtractAttributeValue("colspan"), out colSpan))
			{
				if (colSpan > 1)
				{
					cellProperties.Append(new GridSpan() { Val = colSpan });
				}
			}
			
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
