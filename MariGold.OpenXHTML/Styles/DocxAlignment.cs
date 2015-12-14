namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class DocxAlignment
	{
		private const string textAlign = "text-align";
		
		internal static bool ApplyTextAlign(string styleName, string value, OpenXmlElement styleElement)
		{
			if (string.Compare(textAlign, styleName, StringComparison.InvariantCultureIgnoreCase) != 0) 
			{
				return false;
			}
			
			JustificationValues alignment;
					
			if (DocxAlignment.GetJustificationValue(value, out alignment)) 
			{
				styleElement.Append(new Justification() { Val = alignment });
			}
			
			return true;
		}
		
		internal static bool GetJustificationValue(string value, out JustificationValues alignment)
		{
			alignment = JustificationValues.Left;
			bool assigned = false;

			switch (value.ToLower()) 
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
	}
}
