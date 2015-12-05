namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class StyleParser
	{
		private readonly DocxColor docxColor;
		
		internal const string backGroundColor = "background-color";
		internal const string color = "color";
		internal const string fontFamily = "font-family";
		
		internal StyleParser()
		{
			docxColor = new DocxColor();
		}
		
		internal Shading GetBackGroundColor(string value)
		{
			string hex = docxColor.GetHexColor(value);
					
			if (!string.IsNullOrEmpty(hex))
			{
				return new Shading(){ Fill = hex };
			}
			
			return null;
		}
		
		internal Color GetColor(string value)
		{
			string hex = docxColor.GetHexColor(value);
					
			if (!string.IsNullOrEmpty(hex))
			{
				return new Color(){ Val = hex };
			}
			
			return null;
		}
		
		internal RunFonts GetFonts(string value)
		{
			return new RunFonts() { Ascii = value };
		}
	}
}
