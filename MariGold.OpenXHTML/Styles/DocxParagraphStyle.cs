namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxParagraphStyle
	{
		private bool CheckAlignment(KeyValuePair<string,string> style, ParagraphProperties properties)
		{
			if (DocxAlignment.ApplyTextAlign(style.Key, style.Value, properties)) 
			{
				return true;
			}
			
			return false;
		}
		
		private bool CheckColor(KeyValuePair<string,string> style, ParagraphProperties properties)
		{
			if (DocxColor.ApplyBackGroundColor(style.Key, style.Value, properties)) 
			{
				return true;
			}
			
			return false;
		}
		
		internal void Process(Paragraph element, Dictionary<string, string> styles)
		{
			ParagraphProperties properties = element.ParagraphProperties;
			
			if (element.ParagraphProperties == null) 
			{
				properties = new ParagraphProperties();
			}
			
			foreach (KeyValuePair<string,string> style in styles) 
			{
				if (CheckAlignment(style, properties)) 
				{
					continue;
				}
				
				CheckColor(style, properties);
			}
			
			if (element.ParagraphProperties == null && properties.HasChildren) 
			{
				element.ParagraphProperties = properties;
			}
		}
	}
}
