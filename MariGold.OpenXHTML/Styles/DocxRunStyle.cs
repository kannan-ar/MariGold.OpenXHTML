namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal sealed class DocxRunStyle
	{
		private void CheckFonts(DocxNode docxNode, RunProperties properties)
		{
			string fontFamily = docxNode.ExtractStyleValue(DocxFont.fontFamily);
			string fontWeight = docxNode.ExtractStyleValue(DocxFont.fontWeight);
			string textDecoration=docxNode.ExtractStyleValue(DocxFont.textDecoration);
			
			if (!string.IsNullOrEmpty(fontFamily))
			{
				DocxFont.ApplyFontFamily(fontFamily, properties);
			}
			
			if (!string.IsNullOrEmpty(fontWeight))
			{
				DocxFont.ApplyFontWeight(fontWeight, properties);
			}
				
			if (!string.IsNullOrEmpty(textDecoration))
			{
				DocxFont.ApplyTextDecoration(textDecoration, properties);
			}
		}
		
		private void CheckColor(DocxNode docxNode, RunProperties properties)
		{
			string backgroundColor = docxNode.ExtractStyleValue(DocxColor.backGroundColor);
			string color = docxNode.ExtractStyleValue(DocxColor.color);
			
			if (!string.IsNullOrEmpty(backgroundColor))
			{
				DocxColor.ApplyBackGroundColor(backgroundColor, properties);
			}
			
			if (!string.IsNullOrEmpty(color))
			{
				DocxColor.ApplyColor(color, properties);
			}
		}
		
		internal void Process(Run element, IHtmlNode node)
		{
			RunProperties properties = element.RunProperties;
			DocxNode docxNode = new DocxNode(node);
			
			if (properties == null)
			{
				properties = new RunProperties();
			}
			
			CheckFonts(docxNode, properties);
			CheckColor(docxNode, properties);
			
			if (element.RunProperties == null && properties.HasChildren)
			{
				element.RunProperties = properties;
			}
		}
	}
}
