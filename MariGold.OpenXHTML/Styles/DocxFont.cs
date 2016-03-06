namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class DocxFont
	{
		internal const string fontFamily = "font-family";
		internal const string fontWeight = "font-weight";
		internal const string fontStyle = "font-style";
		internal const string textDecoration = "text-decoration";
		internal const string fontSize = "font-size";
		
		internal const string bold = "bold";
		internal const string bolder = "bolder";
		internal const string italic = "italic";
		internal const string oblique = "oblique";
		internal const string underLine = "underline";
		internal const string lineThrough = "line-through";
		
		internal static void ApplyFontFamily(string style, OpenXmlElement styleElement)
		{
			styleElement.Append(new RunFonts() { Ascii = style });
		}
		
		internal static void ApplyFontWeight(string style, OpenXmlElement styleElement)
		{
			if (string.Compare(bold, style, StringComparison.InvariantCultureIgnoreCase) == 0 ||
			    string.Compare(bolder, style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				styleElement.Append(new Bold());
			}
		}
		
		internal static void ApplyFontItalic(string style, OpenXmlElement styleElement)
		{
			if (string.Compare(italic, style, StringComparison.InvariantCultureIgnoreCase) == 0 &&
			    string.Compare(oblique, style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				styleElement.Append(new Italic());
			}
		}
		
		internal static void ApplyTextDecoration(string style, OpenXmlElement styleElement)
		{
			if (string.Compare(style, underLine, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				styleElement.Append(new Underline(){ Val = UnderlineValues.Single });
			}
			else
			if (string.Compare(style, lineThrough, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				styleElement.Append(new Strike());
			}
		}
		
		internal static void ApplyUnderline(OpenXmlElement styleElement)
		{
			styleElement.Append(new Underline(){ Val = UnderlineValues.Single });
		}
		
		internal static void ApplyFontItalic(OpenXmlElement styleElement)
		{
			styleElement.Append(new Italic());
		}
		
		internal static void ApplyBold(OpenXmlElement styleElement)
		{
			styleElement.Append(new Bold());
		}
		
		internal static void ApplyFontSize(string style, OpenXmlElement styleElement)
		{
			int fontSize = DocxUnits.HalfPointFromStyle(style);
			
			if (fontSize != 0)
			{
				styleElement.Append(new FontSize(){ Val = fontSize.ToString() });
			}
		}
		
		internal static void ApplyFont(int size, bool isBold, OpenXmlElement styleElement)
		{
			FontSize fontSize = new FontSize(){ Val = size.ToString() };
			
			if (isBold)
			{
				styleElement.Append(new Bold());
			}
			
			styleElement.Append(fontSize);
		}
	}
}
