namespace MariGold.OpenXHTML
{
	using System;
	using System.Text.RegularExpressions;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class DocxBorder
	{
		private readonly static Regex borderWidth = new Regex("\\d+((px)|(pt)|(cm)|(em))");
		private readonly static Regex intMatch = new Regex("\\d+");
		private readonly static string[] boderStyles = { "dotted", "dashed", "solid", "double", "groove", "ridge", "inset", "outset", "none", "hidden" };
		
		private static string GetBorderWidth(ref string borderStyle)
		{
			string width = string.Empty;
			borderWidth.Options = RegexOptions.CultureInvariant | RegexOptions.IgnoreCase;
			
			Match match = borderWidth.Match(borderStyle);
			
			if (match.Success)
			{
				Match intValue = intMatch.Match(match.Value);
				
				if (intValue.Success)
				{
					width = intValue.Value;
				}
				
				borderStyle = borderStyle.Replace(match.Value, string.Empty);
			}
			
			return width;
		}
		
		private static string GetBorderStyle(ref string borderStyle)
		{
			foreach (string style in boderStyles)
			{
				if (borderStyle.Contains(style))
				{
					int index = borderStyle.IndexOf(style, StringComparison.InvariantCultureIgnoreCase);
					
					if (index != -1)
					{
						//Length may vary in different culture?
						borderStyle = borderStyle.Remove(index, style.Length);
						
						return style;
					}
				}
			}
			
			return string.Empty;
		}
		
		internal static void ApplyDefaultBorders(OpenXmlCompositeElement element)
		{
			TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
			LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
			BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
			RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
				
			element.Append(topBorder);
			element.Append(leftBorder);
			element.Append(bottomBorder);
			element.Append(rightBorder);
		}
		
		
	}
}
