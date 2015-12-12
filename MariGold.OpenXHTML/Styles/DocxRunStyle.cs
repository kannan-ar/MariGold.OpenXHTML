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
			List<OpenXmlElement> list = new List<OpenXmlElement>();
			
			foreach (KeyValuePair<string,string> style in styles)
			{
				if (DocxColor.IsBackGroundColor(style.Key))
				{
					Shading shading = DocxColor.GetBackGroundColor(style.Value);
					
					if (shading != null)
					{
						list.Add(shading);
					}
					
					continue;
				}
				
				if (DocxColor.IsColor(style.Key))
				{
					Color color = DocxColor.GetColor(style.Value);
					
					if (color != null)
					{
						list.Add(color);
					}
					
					continue;
				}
				
				if (DocxFont.IsFontFamily(style.Key))
				{
					RunFonts font = DocxFont.GetFonts(style.Value);
					
					if (font != null)
					{
						list.Add(font);
					}
					
					continue;
				}
			}
			
			if (list.Count > 0)
			{
				if (element.RunProperties == null)
				{
					element.RunProperties = new RunProperties();
				}
				
				foreach (OpenXmlElement item in list)
				{
					element.RunProperties.Append(item);
				}
			}
		}
	}
}
