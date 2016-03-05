namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal sealed class DocxParagraphStyle
	{
		private void ProcessBorder(DocxNode docxNode, ParagraphProperties properties)
		{
			ParagraphBorders paragraphBorders = new ParagraphBorders();
			
			DocxBorder.ApplyBorders(paragraphBorders,
				docxNode.ExtractStyleValue(DocxBorder.borderName),
				docxNode.ExtractStyleValue(DocxBorder.leftBorderName),
				docxNode.ExtractStyleValue(DocxBorder.topBorderName),
				docxNode.ExtractStyleValue(DocxBorder.rightBorderName),
				docxNode.ExtractStyleValue(DocxBorder.bottomBorderName),
				false);
			
			if (paragraphBorders.HasChildren)
			{
				properties.Append(paragraphBorders);
			}
		}
		
		internal void Process(Paragraph element, IHtmlNode node)
		{
			ParagraphProperties properties = element.ParagraphProperties;
			DocxNode docxNode = new DocxNode(node);
			
			if (properties == null)
			{
				properties = new ParagraphProperties();
			}
			
			string textAlign = docxNode.ExtractStyleValue(DocxAlignment.textAlign);
			
			if (!string.IsNullOrEmpty(textAlign))
			{
				DocxAlignment.ApplyTextAlign(textAlign, properties);
			}
			
			string backgroundColor = docxNode.ExtractStyleValue(DocxColor.backGroundColor);
			
			if (!string.IsNullOrEmpty(backgroundColor))
			{
				DocxColor.ApplyBackGroundColor(backgroundColor, properties);
			}
			
			ProcessBorder(docxNode, properties);
			
			if (element.ParagraphProperties == null && properties.HasChildren)
			{
				element.ParagraphProperties = properties;
			}
		}
	}
}
