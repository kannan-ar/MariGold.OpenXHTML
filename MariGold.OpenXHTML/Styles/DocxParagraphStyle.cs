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
			ParagraphProperties properties = element.ParagraphProperties;
			
			if (element.ParagraphProperties == null) 
			{
				properties = new ParagraphProperties();
			}
			
			foreach (KeyValuePair<string,string> style in styles) 
			{
				if (DocxAlignment.IsTextAlign(style.Key)) 
				{
					JustificationValues alignment;
					
					if (DocxAlignment.GetJustificationValue(style.Value, out alignment)) 
					{
						properties.Append(new Justification() { Val = alignment });
					}
					
					continue;
				}
				
				if (DocxColor.IsBackGroundColor(style.Key)) 
				{
					Shading shading = DocxColor.GetBackGroundColor(style.Value);
					
					if (shading != null) 
					{
						properties.Append(shading);
					}
					
					continue;
				}
			}
			
			if (element.ParagraphProperties == null && properties.HasChildren) 
			{
				element.ParagraphProperties = properties;
			}
		}
	}
}
