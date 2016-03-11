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
			string fontStyle = docxNode.ExtractStyleValue(DocxFont.fontStyle);
			
			if (!string.IsNullOrEmpty(fontFamily))
			{
				DocxFont.ApplyFontFamily(fontFamily, properties);
			}
			
			if (!string.IsNullOrEmpty(fontWeight))
			{
				DocxFont.ApplyFontWeight(fontWeight, properties);
			}
				
			if (!string.IsNullOrEmpty(fontStyle))
			{
				DocxFont.ApplyFontStyle(fontStyle, properties);
			}
		}
		
		private void CheckFontStyle(DocxNode docxNode, RunProperties properties)
		{
			string fontSize = docxNode.ExtractStyleValue(DocxFont.fontSize);
			string textDecoration = docxNode.ExtractStyleValue(DocxFont.textDecoration);
			
			if (!string.IsNullOrEmpty(fontSize))
			{
				DocxFont.ApplyFontSize(fontSize, properties);
			}
			
			if (!string.IsNullOrEmpty(textDecoration))
			{
				DocxFont.ApplyTextDecoration(textDecoration, properties);
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
			
			//Order of assigning styles to run property is important. The order should not change.
			CheckFonts(docxNode, properties);
			
			string color = docxNode.ExtractStyleValue(DocxColor.color);
			
			if (!string.IsNullOrEmpty(color))
			{
				DocxColor.ApplyColor(color, properties);
			}
			
			CheckFontStyle(docxNode, properties);
			
			string backgroundColor = docxNode.ExtractStyleValue(DocxColor.backGroundColor);
			
			if (!string.IsNullOrEmpty(backgroundColor))
			{
				DocxColor.ApplyBackGroundColor(backgroundColor, properties);
			}
			
			if (element.RunProperties == null && properties.HasChildren)
			{
				element.RunProperties = properties;
			}
		}
	}
}
