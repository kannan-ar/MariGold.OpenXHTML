namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal sealed class DocxTableStyle
	{
		private void ApplyTableBorder(TableProperties tableProperties, IHtmlNode node)
		{
			DocxNode docxNode = new DocxNode(node);
			
			string borderStyle = docxNode.ExtractAttributeValue("border");
			
			if (borderStyle == "1")
			{
				TableBorders tableBorders = new TableBorders();
				DocxBorder.ApplyDefaultBorders(tableBorders);
				tableProperties.Append(tableBorders);
			}
			else
			{
				borderStyle = docxNode.ExtractStyleValue("border");
				
				if (!string.IsNullOrEmpty(borderStyle))
				{
					TableBorders tableBorders = new TableBorders();
					DocxBorder.ApplyBorders(tableBorders, borderStyle, null, null, null, null);
					tableProperties.Append(tableBorders);
					
				}
			}
		}
		
		internal void Process(TableProperties tableProperties, IHtmlNode node)
		{
			ApplyTableBorder(tableProperties, node);
		}
	}
}
