namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	
	internal static class DocxUnits
	{
		internal static Int16 GetDxaFromPixel(Int16 pixel)
		{
			return (Int16)(pixel * 20);
		}
	}
}
