namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;

    internal static class DocxUnits
    {
        private static Regex Digit => new Regex("(\\.)?\\d+(\\.?\\d+)?");

        private static Dictionary<string, decimal> ToPt => new Dictionary<string, decimal>
        {
            { "px", 1 },
            { "pt", 1 },
            { "em", 12 },
            { "cm", 28.34m },
            { "in", 72 }
        };

        private static Dictionary<string, decimal> ToPx => new Dictionary<string, decimal>
        {
            { "px", 1 },
            { "pt", .75m },
            { "em", 12 },
            { "cm", 28.34m },
            { "in", 72 }
        };

        private static Dictionary<string, decimal> NamedFontSizeList => new Dictionary<string, decimal>
        {
            { "xx-small", 8 },
            { "x-small", 9 },
            { "smaller", 10 },
            { "small", 11 },
            { "medium", 12 },
            { "large", 13 },
            { "larger", 14 },
            { "x-large", 20 },
            { "xx-large", 24 }
        };

        internal const string width = "width";

        private static bool TryGetValue(string style, out decimal value)
        {
            value = 0;

            if (string.IsNullOrEmpty(style))
            {
                return false;
            }

            Match match = Digit.Match(style);

            if (!match.Success)
            {
                return false;
            }

            if (!decimal.TryParse(match.Value, out value))
            {
                return false;
            }

            return true;
        }

        private static bool ConvertToPt(string style, out decimal value)
        {
            if (!TryGetValue(style, out value))
            {
                return false;
            }

            //Value is on percentage. So no need to convert to pt. just returning after value extraction.
            if (style.Contains("%"))
            {
                return true;
            }

            foreach (var item in ToPt)
            {
                if (style.IndexOf(item.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    value *= item.Value;
                    return true;
                }
            }

            return false;
        }

        private static decimal ConvertPercentageToPt(decimal value)
        {
            return (value / 100) * DocxFontStyle.defaultFontSizeInPixel;
        }

        private static bool ExtractNamedFontSize(string style, out decimal pt)
        {
            style = style.Trim();

            var item = NamedFontSizeList.FirstOrDefault(x => string.Compare(x.Key, style.Trim(), StringComparison.InvariantCultureIgnoreCase) == 0);
            pt = item.Value;
            return !item.Equals(default(KeyValuePair<string, decimal>));
        }

        internal static Int32 GetDxaFromPixel(Int32 pixel)
        {
            return pixel * 20;
        }

        internal static bool ConvertToPx(string style, out decimal value)
        {
            if (!TryGetValue(style, out value))
            {
                return false;
            }

            //Percentage is not supported
            if (style.Contains("%"))
            {
                return false;
            }

            foreach (var item in ToPx)
            {
                if (style.IndexOf(item.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    value *= item.Value;
                    value = decimal.Round(value);
                    return true;
                }
            }

            return true;
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
                value *= 50;//Convert to fifties
                unit = TableWidthUnitValues.Pct;

                return true;
            }
            else
            {
                value *= 20; //Convert to Twentieths of a point
                unit = TableWidthUnitValues.Dxa;

                return true;
            }
        }

        internal static decimal HalfPointFromStyle(string style)
        {
            if (ExtractNamedFontSize(style, out decimal pt))
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
                pt *= 2;
            }

            return pt;
        }

        internal static decimal GetDxaFromStyle(string style)
        {
            if (ConvertToPt(style, out decimal value))
            {
                if (style.Contains("%"))
                {
                    value = ConvertPercentageToPt(value);
                }

                return value * 20;
            }

            return -1;
        }

        internal static decimal GetDxaFromNumber(decimal number)
        {
            return number * DocxFontStyle.defaultFontSizeInPixel * 20;
        }
    }
}
