namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	
	internal static class DocxUnits
	{
		internal static StringValue GetDxaFromPixel(Int16 pixel)
		{
			return (pixel * 20).ToString();
		}
	}
}
