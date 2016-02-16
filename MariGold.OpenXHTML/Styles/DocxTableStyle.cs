namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal sealed class DocxTableStyle
	{
		internal void Process(TableProperties tableProperties, DocxTableProperties docxProperties, IHtmlNode node)
		{
			DocxNode docxNode = new DocxNode(node);
			
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
			
			if(cellBorders.HasChildren)
			{
				cellProperties.Append(cellBorders);
			}
			
			if(cellProperties.HasChildren)
			{
				cell.Append(cellProperties);
			}
		}
	}
}
