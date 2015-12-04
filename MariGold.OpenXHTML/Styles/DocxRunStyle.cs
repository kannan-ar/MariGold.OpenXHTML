namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxRunStyle : DocxStyle<Run>
	{
		private readonly DocxColor docxColor;
		
		public DocxRunStyle()
		{
			docxColor = new DocxColor();
		}
		
		internal override void Process(Run element, Dictionary<string, string> styles)
		{
			List<OpenXmlElement> list = new List<OpenXmlElement>();
			
			foreach (KeyValuePair<string,string> style in styles)
			{
				if (string.Compare(backGroundColor, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					string hex = docxColor.GetHexColor(style.Value);
					
					if (!string.IsNullOrEmpty(hex))
					{
						list.Add(new Shading(){ Fill = hex });
					}
				}
				else if (string.Compare(color, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					string hex = docxColor.GetHexColor(style.Value);
					
					list.Add(new Color(){ Val = hex });
				}
				else if (string.Compare(fontFamily, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					
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
