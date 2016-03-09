namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal sealed class DocxParagraphStyle
	{
		private const string margin = "margin";
		private const string marginTop = "margin-top";
		private const string marginBottom = "margin-bottom";
		
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
		
		private void ProcessMargin(DocxNode docxNode, ParagraphProperties properties)
		{
			string marginVal = docxNode.ExtractStyleValue(margin);
			string marginTopVal = docxNode.ExtractStyleValue(marginTop);
			string marginBottomVal = docxNode.ExtractStyleValue(marginBottom);
			
			bool hasTopMargin = false;
			bool hasBottomMargin = false;
			
			hasTopMargin = !string.IsNullOrEmpty(marginTopVal);
			hasBottomMargin = !string.IsNullOrEmpty(marginBottomVal);
			
			if (!hasTopMargin && !string.IsNullOrEmpty(marginVal))
			{
				marginTopVal = marginVal;
				hasTopMargin = true;
			}
			
			if (!hasBottomMargin && !string.IsNullOrEmpty(marginVal))
			{
				marginBottomVal = marginVal;
				hasBottomMargin = true;
			}
			
			if (hasTopMargin || hasBottomMargin)
			{
				SpacingBetweenLines spacing = new SpacingBetweenLines();
				
				if (hasTopMargin)
				{
					spacing.Before = DocxUnits.GetDxaFromStyle(marginTopVal).ToString();
				}
				
				if (hasBottomMargin)
				{
					spacing.After = DocxUnits.GetDxaFromStyle(marginBottomVal).ToString();
				}
				
				properties.Append(spacing);
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
			
			ProcessMargin(docxNode, properties);
			
			if (element.ParagraphProperties == null && properties.HasChildren)
			{
				element.ParagraphProperties = properties;
			}
		}
	}
}
