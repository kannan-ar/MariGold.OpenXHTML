namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class DocxFont
	{
		internal const string fontFamily = "font-family";
		internal const string fontWeight = "font-weight";
		internal const string bold="bold";
		internal const string bolder="bolder";
		
		internal static bool IsFontFamily(string styleName)
		{
			return string.Compare(fontFamily, styleName, StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal static bool IsFontWeight(string styleName,string value)
		{
			return string.Compare(fontWeight, styleName, StringComparison.InvariantCultureIgnoreCase) == 0 &&
				string.Compare(bold, value, StringComparison.InvariantCultureIgnoreCase) == 0 &&
				string.Compare(bolder, value, StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal static RunFonts GetFonts(string value)
		{
			return new RunFonts() { Ascii = value };
		}
		
		internal static Bold GetBold()
		{
			return new Bold();
		}
	}
}
