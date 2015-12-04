namespace MariGold.OpenXHTML
{
	using System;
	using System.Drawing;
	
	internal sealed class DocxColor
	{
		private bool IsRGB(string styleValue)
		{
			return styleValue.IndexOf("rgb", StringComparison.CurrentCultureIgnoreCase) >= 0;
		}
		
		private string GetHex(string rgb)
		{
			int startIndex = rgb.IndexOf("(");
			int endIndex = rgb.IndexOf(")");
			string hex = string.Empty;
			
			if (startIndex >= 0 && endIndex > startIndex)
			{
				string color = rgb.Substring(startIndex + 1, endIndex - startIndex - 1);
				
				if (!string.IsNullOrEmpty(color))
				{
					string[] colors = color.Split(new char[]{ ',' }, StringSplitOptions.RemoveEmptyEntries);
					
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
		
		public bool IsHex(string styleValue)
		{
			return styleValue.IndexOf("#") >= 0;
		}
		
		public string GetHexColor(string styleValue)
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
			else if(IsHex(styleValue))
			{
				hex = styleValue.Replace("#", string.Empty);
			}
			
			return hex;
		}
	}
}
