namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxRunStyle : DocxStyle<Run>
	{
		internal override void Process(Run element, Dictionary<string, string> styles)
		{
			RunProperties properties = element.RunProperties;
			
			if (properties == null) 
			{
				properties = new RunProperties();
			}
			
			foreach (KeyValuePair<string,string> style in styles) 
			{
				if (DocxColor.IsBackGroundColor(style.Key)) 
				{
					Shading shading = DocxColor.GetBackGroundColor(style.Value);
					
					if (shading != null) 
					{
						properties.Append(shading);
					}
					
					continue;
				}
				
				if (DocxColor.IsColor(style.Key)) 
				{
					Color color = DocxColor.GetColor(style.Value);
					
					if (color != null) 
					{
						properties.Append(color);
					}
					
					continue;
				}
				
				if (DocxFont.IsFontFamily(style.Key)) 
				{
					RunFonts font = DocxFont.GetFonts(style.Value);
					
					if (font != null) 
					{
						properties.Append(font);
					}
					
					continue;
				}
				
				if (DocxFont.IsFontWeight(style.Key, style.Value))
				{
					properties.Append(DocxFont.GetBold());
					
					continue;
				}
			}
			
			if (element.RunProperties == null && properties.HasChildren) 
			{
				element.RunProperties = properties;
			}
		}
	}
}
