namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxRunStyle : DocxStyle<Run>
	{
		private bool CheckFonts(KeyValuePair<string,string> style, RunProperties properties)
		{
			if(DocxFont.ApplyFontFamily(style.Key,style.Value, properties))
			{
				return true;
			}
			
			if(DocxFont.ApplyFontWeight(style.Key,style.Value,properties))
			{
				return true;
			}
				
			if(DocxFont.ApplyTextDecoration(style.Key, style.Value, properties))
			{
				return true;
			}
			
			return false;
		}
		
		private bool CheckColor(KeyValuePair<string,string> style, RunProperties properties)
		{
			if (DocxColor.ApplyBackGroundColor(style.Key,style.Value,properties)) 
			{
				return true;
			}
				
			if (DocxColor.ApplyColor(style.Key,style.Value,properties)) 
			{
				return true;
			}
			
			return false;
		}
		
		internal override void Process(Run element, Dictionary<string, string> styles)
		{
			RunProperties properties = element.RunProperties;
			
			if (properties == null) 
			{
				properties = new RunProperties();
			}
			
			foreach (KeyValuePair<string,string> style in styles) 
			{
				if(CheckColor(style, properties))
				{
					continue;
				}
				
				CheckFonts(style, properties);
			}
			
			if (element.RunProperties == null && properties.HasChildren) 
			{
				element.RunProperties = properties;
			}
		}
	}
}
