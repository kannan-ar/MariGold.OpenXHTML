namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class DocxAlignment
	{
		internal const string textAlign = "text-align";
		
		internal static void ApplyTextAlign(string style, OpenXmlElement styleElement)
		{
			JustificationValues alignment;
					
			if (DocxAlignment.GetJustificationValue(style, out alignment)) 
			{
				styleElement.Append(new Justification() { Val = alignment });
			}
		}
		
		internal static bool GetJustificationValue(string style, out JustificationValues alignment)
		{
			alignment = JustificationValues.Left;
			bool assigned = false;

			switch (style.ToLower()) 
			{
				case "right":
					assigned = true;
					alignment = JustificationValues.Right;
					break;

				case "left":
					assigned = true;
					alignment = JustificationValues.Left;
					break;

				case "center":
					assigned = true;
					alignment = JustificationValues.Center;
					break;
			}

			return assigned;
		}
		
		internal static void AlignCenter(OpenXmlElement styleElement)
		{
			styleElement.Append(new Justification() { Val = JustificationValues.Center });
		}
	}
}
