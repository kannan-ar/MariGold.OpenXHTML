namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxRunStyle : DocxStyle<Run>
	{
		private readonly StyleParser parser;
		
		public DocxRunStyle()
		{
			parser = new StyleParser();
		}
		
		internal override void Process(Run element, Dictionary<string, string> styles)
		{
			List<OpenXmlElement> list = new List<OpenXmlElement>();
			
			foreach (KeyValuePair<string,string> style in styles)
			{
				if (string.Compare(StyleParser.backGroundColor, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					Shading shading = parser.GetBackGroundColor(style.Value);
					
					if (shading != null)
					{
						list.Add(shading);
					}
					
					continue;
				}
				
				if (string.Compare(StyleParser.color, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					Color color = parser.GetColor(style.Value);
					
					if (color != null)
					{
						list.Add(color);
					}
					
					continue;
				}
				
				if (string.Compare(StyleParser.fontFamily, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					RunFonts font = parser.GetFonts(style.Value);
					
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
