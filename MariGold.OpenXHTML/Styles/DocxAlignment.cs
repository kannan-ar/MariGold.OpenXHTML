namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class DocxAlignment
	{
		internal const string textAlign = "text-align";
		
		internal static bool IsTextAlign(string styleName)
		{
			return string.Compare(textAlign, styleName, StringComparison.InvariantCultureIgnoreCase) == 0;
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
