namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.Text.RegularExpressions;
	using System.Collections.Generic;
	
	internal static class DocxUnits
	{
		private static Regex digit;
		private static Dictionary<string,decimal> toPt;
		
		internal const string width = "width";
		
		private static bool ConvertToPt(string style, out decimal value)
		{
			value = 0;
			
			if (string.IsNullOrEmpty(style))
			{
				return false;
			}
			
			Match match = digit.Match(style);
			
			if (!match.Success)
			{
				return false;
			}
			
			if (!decimal.TryParse(match.Value, out value))
			{
				return false;
			}
			
			//Value is on percentage. So no need to convert to pt. just returning after value extraction.
			if (style.Contains("%"))
			{
				return true;
			}
			
			foreach (var item in toPt)
			{
				if (style.IndexOf(item.Key, StringComparison.OrdinalIgnoreCase) >= 0)
				{
					value = value * item.Value;
					return true;
				}
			}
			
			return false;
		}
		
		private static decimal ConvertPercentageToPt(decimal value)
		{
			return value * .12m;
		}
		
		private static bool ExtractNamedFontSize(string style, out decimal pt)
		{
			pt = 0;
			
			style = style.Trim();
			
			if (string.Compare("xx-small", style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				pt = 8;
				return true;
			}
			else
			if (string.Compare("x-small", style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				pt = 9;
				return true;
			}
			else
			if (string.Compare("smaller", style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				pt = 10;
				return true;
			}
			else
			if (string.Compare("small", style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				pt = 11;
				return true;
			}
			else
			if (string.Compare("medium", style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				pt = 12;
				return true;
			}
			else
			if (string.Compare("large", style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				pt = 13;
				return true;
			}
			else
			if (string.Compare("larger", style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				pt = 14;
				return true;
			}
			else
			if (string.Compare("x-large", style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				pt = 20;
				return true;
			}
			else
			if (string.Compare("xx-large", style, StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				pt = 24;
				return true;
			}
			
			return false;
		}
		
		static DocxUnits()
		{
            digit = new Regex("(\\.)?\\d+(\\.?\\d+)?");
			
			toPt = new Dictionary<string, decimal>();
			toPt.Add("px", 1);
			toPt.Add("pt", 1);
			toPt.Add("em", 12); 
			toPt.Add("cm", 28.34m);
			toPt.Add("in", 72);
		}
		
		internal static Int32 GetDxaFromPixel(Int32 pixel)
		{
			return pixel * 20;
		}
		
		internal static bool TableUnitsFromStyle(string style, out decimal value, out TableWidthUnitValues unit)
		{
			unit = TableWidthUnitValues.Nil;
			
			if (!ConvertToPt(style, out value))
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
				value = value * 20; //Convert to Twentieths of a point
				unit = TableWidthUnitValues.Dxa;
								
				return true;
			}
		}
		
		internal static decimal HalfPointFromStyle(string style)
		{
			decimal pt = 0;
			
			if (ExtractNamedFontSize(style, out pt))
			{
				return pt * 2;
			}
			
			if (!ConvertToPt(style, out pt))
			{
				return 0;
			}
			
			if (style.Contains("%"))
			{
				pt = ConvertPercentageToPt(pt) * 2;
			}
			else
			{
				pt = pt * 2;
			}
			
			return pt;
		}
		
		internal static decimal GetDxaFromStyle(string style)
		{
			decimal value;
			
			if (ConvertToPt(style, out value))
			{
                if(style.Contains("%"))
                {
                    value = ConvertPercentageToPt(value);
                }

				return value * 20;
			}
			
			return -1;
		}

        internal static Int32 GetDxaFromNumber(Int32 number)
        {
            return number * DocxFont.defaultFontSizeInPixel * 20;
        }
	}
}
