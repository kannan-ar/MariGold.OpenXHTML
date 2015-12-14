namespace MariGold.OpenXHTML
{
	using System;
	using System.Drawing;
	using DocumentFormat.OpenXml;
	using Word = DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class DocxColor
	{
		private const string backGroundColor = "background-color";
		private const string color = "color";
		
		private static bool IsRGB(string styleValue)
		{
			return styleValue.IndexOf("rgb", StringComparison.CurrentCultureIgnoreCase) >= 0;
		}
		
		private static string GetHex(string rgb)
		{
			int startIndex = rgb.IndexOf("(");
			int endIndex = rgb.IndexOf(")");
			string hex = string.Empty;
			
			if (startIndex >= 0 && endIndex > startIndex) 
			{
				string val = rgb.Substring(startIndex + 1, endIndex - startIndex - 1);
				
				if (!string.IsNullOrEmpty(val)) 
				{
					string[] colors = val.Split(new char[]{ ',' }, StringSplitOptions.RemoveEmptyEntries);
					
					if (colors.Length > 2) 
					{
						int r, g, b = 0;
						
						r = Convert.ToInt32(colors[0]);
						g = Convert.ToInt32(colors[1]);
						b = Convert.ToInt32(colors[2]);
						
						Color c = Color.FromArgb(r, g, b);
						
						hex = c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
					}
				}
			}
			
			return hex;
		}
		
		internal static bool IsHex(string styleValue)
		{
			return styleValue.IndexOf("#") >= 0;
		}
		
		internal static string GetHexColor(string styleValue)
		{
			string hex = string.Empty;
			
			if (string.IsNullOrEmpty(styleValue)) 
			{
				return string.Empty;
			}
			
			if (IsRGB(styleValue)) 
			{
				hex = GetHex(styleValue);
			}
			else if (IsHex(styleValue)) 
			{
				hex = styleValue.Replace("#", string.Empty);
			}
			
			return hex;
		}
		
		internal static bool ApplyBackGroundColor(string styleName, string value, OpenXmlElement styleElement)
		{
			if (string.Compare(backGroundColor, styleName, StringComparison.InvariantCultureIgnoreCase) == 0) 
			{
				string hex = GetHexColor(value);
					
				if (!string.IsNullOrEmpty(hex)) 
				{
					styleElement.Append(new Word.Shading(){ Fill = hex });
				}
				
				return true;
			}
			
			return false;
		}
		
		internal static bool ApplyColor(string styleName, string value, OpenXmlElement styleElement)
		{
			if (string.Compare(color, styleName, StringComparison.InvariantCultureIgnoreCase) == 0) 
			{
				string hex = GetHexColor(value);
					
				if (!string.IsNullOrEmpty(hex)) 
				{
					styleElement.Append(new Word.Color(){ Val = hex });
				}
				
				return true;
			}
			
			return false;
		}
	}
}
