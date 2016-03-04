namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.Text.RegularExpressions;
	using System.Collections.Generic;
	
	internal static class DocxUnits
	{
		private static Regex digit;
		private static Dictionary<string,int> dxaUnits;
		
		internal const string width = "width";
		
		static DocxUnits()
		{
			digit = new Regex("\\d+");
			
			dxaUnits = new Dictionary<string, int>();
			dxaUnits.Add("px", 20);
			dxaUnits.Add("pt", 20); //
			dxaUnits.Add("em", 320); // 16 * 20 to convert to pixels. Assuming 16 is the default pixel size. Need to rework
			dxaUnits.Add("cm", 567);
			dxaUnits.Add("in", 1440);
		}
		
		internal static Int16 GetDxaFromPixel(Int16 pixel)
		{
			return (Int16)(pixel * 20);
		}
		
		internal static bool TableUnitsFromStyle(string style, out int value, out TableWidthUnitValues unit)
		{
			value = 0;
			unit = TableWidthUnitValues.Nil;
			
			if (string.IsNullOrEmpty(style))
			{
				return false;
			}
			
			Match match = digit.Match(style);
			
			if (!match.Success)
			{
				return false;
			}
			
			if (!Int32.TryParse(match.Value, out value))
			{
				return false;
			}
			
			if (style.Contains("%"))
			{
				value = value * 50;//Convert to fifties
				unit = TableWidthUnitValues.Pct;
						
				return true;
			}
			else
			{
				foreach (var item in dxaUnits)
				{
					if (style.IndexOf(item.Key, StringComparison.InvariantCultureIgnoreCase) >= 0)
					{
						value = value * item.Value;
						unit = TableWidthUnitValues.Dxa;
								
						return true;
					}
				}
			}
			
			return false;
		}
	}
}
