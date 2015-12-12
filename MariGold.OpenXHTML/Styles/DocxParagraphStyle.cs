namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxParagraphStyle : DocxStyle<Paragraph>
	{
		internal override void Process(Paragraph element, Dictionary<string, string> styles)
		{
			List<OpenXmlElement> list = new List<OpenXmlElement>();
			
			foreach (KeyValuePair<string,string> style in styles)
			{
				if (DocxAlignment.IsTextAlign(style.Key))
				{
					JustificationValues alignment;
					
					if (DocxAlignment.GetJustificationValue(style.Value, out alignment))
					{
						list.Add(new Justification() { Val = alignment });
					}
					
					continue;
				}
				
				if (DocxColor.IsBackGroundColor(style.Key))
				{
					Shading shading = DocxColor.GetBackGroundColor(style.Value);
					
					if (shading != null)
					{
						list.Add(shading);
					}
					
					continue;
				}
			}
			
			if (list.Count > 0)
			{
				if (element.ParagraphProperties == null)
				{
					element.ParagraphProperties = new ParagraphProperties();
				}
				
				foreach (OpenXmlElement item in list)
				{
					element.ParagraphProperties.Append(item);
				}
			}
		}
	}
}
