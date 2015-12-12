namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class DocxFont
	{
		internal const string fontFamily = "font-family";
		
		internal static bool IsFontFamily(string styleName)
		{
			return string.Compare(fontFamily, styleName, StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal static RunFonts GetFonts(string value)
		{
			return new RunFonts() { Ascii = value };
		}
	}
}
