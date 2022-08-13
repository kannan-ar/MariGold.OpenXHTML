namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml;
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    internal static class DocxColor
    {
        private static IDictionary<string, string> namedColors;

        internal const string backGroundColor = "background-color";
        internal const string backGround = "background";
        internal const string color = "color";

        static DocxColor()
        {
            SetNamedColors();
        }

        private static bool IsRGB(string styleValue)
        {
            return styleValue.IndexOf("rgb", StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        private static string GetHex(string rgb)
        {
            int startIndex = rgb.IndexOf("(");
            int endIndex = rgb.IndexOf(")");
            string hex = string.Empty;

            if (startIndex >= 0 && endIndex > startIndex)
            {
                string val = rgb.Substring(startIndex + 1, endIndex - startIndex - 1);

                if (!string.IsNullOrEmpty(val))
                {
                    string[] colors = val.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    if (colors.Length > 2)
                    {
                        int r, g;

                        r = Convert.ToInt32(colors[0]);
                        g = Convert.ToInt32(colors[1]);
                        int b = Convert.ToInt32(colors[2]);

                        Color c = Color.FromArgb(r, g, b);

                        hex = c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
                    }
                }
            }

            return hex;
        }

        private static void SetNamedColors()
        {
            //Taken from https://en.wikipedia.org/wiki/Web_colors#X11_color_names

            namedColors = new Dictionary<string, string>
            {
                { "pink", "FFC0CB" },
                { "lightpink", "FFB6C1" },
                { "hotpink", "FF69B4" },
                { "deeppink", "FF1493" },
                { "palevioletred", "DB7093" },
                { "mediumvioletred", "C71585" },
                { "lightsalmon", "FFA07A" },
                { "salmon", "FA8072" },
                { "darksalmon", "E9967A" },
                { "lightcoral", "F08080" },
                { "indianred", "CD5C5C" },
                { "crimson", "DC143C" },
                { "firebrick", "B22222" },
                { "darkred", "8B0000" },
                { "red", "FF0000" },
                { "orangered", "FF4500" },
                { "tomato", "FF6347" },
                { "coral", "FF7F50" },
                { "darkorange", "FF8C00" },
                { "orange", "FFA500" },
                { "yellow", "FFFF00" },
                { "lightyellow", "FFFFE0" },
                { "lemonchiffon", "FFFACD" },
                { "lightgoldenrodyellow", "FAFAD2" },
                { "papayawhip", "FFEFD5" },
                { "moccasin", "FFE4B5" },
                { "peachpuff", "FFDAB9" },
                { "palegoldenrod", "EEE8AA" },
                { "khaki", "F0E68C" },
                { "darkkhaki", "BDB76B" },
                { "gold", "FFD700" },
                { "cornsilk", "FFF8DC" },
                { "blanchedalmond", "FFEBCD" },
                { "bisque", "FFE4C4" },
                { "navajowhite", "FFDEAD" },
                { "wheat", "F5DEB3" },
                { "burlywood", "DEB887" },
                { "tan", "D2B48C" },
                { "rosybrown", "BC8F8F" },
                { "sandybrown", "F4A460" },
                { "goldenrod", "DAA520" },
                { "darkgoldenrod", "B8860B" },
                { "peru", "CD853F" },
                { "chocolate", "D2691E" },
                { "saddlebrown", "8B4513" },
                { "sienna", "A0522D" },
                { "brown", "A52A2A" },
                { "maroon", "800000" },
                { "darkolivegreen", "556B2F" },
                { "olive", "808000" },
                { "olivedrab", "6B8E23" },
                { "yellowgreen", "9ACD32" },
                { "limegreen", "32CD32" },
                { "lime", "00FF00" },
                { "lawngreen", "7CFC00" },
                { "chartreuse", "7FFF00" },
                { "greenyellow", "ADFF2F" },
                { "springgreen", "00FF7F" },
                { "mediumspringgreen", "00FA9A" },
                { "lightgreen", "90EE90" },
                { "palegreen", "98FB98" },
                { "darkseagreen", "8FBC8F" },
                { "mediumaquamarine", "66CDAA" },
                { "mediumseagreen", "3CB371" },
                { "seagreen", "2E8B57" },
                { "forestgreen", "228B22" },
                { "green", "008000" },
                { "darkgreen", "006400" },
                //Cyan colors
                { "aqua", "00FFFF" },
                { "cyan", "00FFFF" },
                { "lightcyan", "E0FFFF" },
                { "paleturquoise", "AFEEEE" },
                { "aquamarine", "7FFFD4" },
                { "turquoise", "40E0D0" },
                { "mediumturquoise", "48D1CC" },
                { "darkturquoise", "00CED1" },
                { "lightseagreen", "20B2AA" },
                { "cadetblue", "5F9EA0" },
                { "darkcyan", "008B8B" },
                { "teal", "008080" },
                //Blue colors
                { "lightsteelblue", "B0C4DE" },
                { "powderblue", "B0E0E6" },
                { "lightblue", "ADD8E6" },
                { "skyblue", "87CEEB" },
                { "lightskyblue", "87CEFA" },
                { "deepskyblue", "00BFFF" },
                { "dodgerblue", "1E90FF" },
                { "cornflowerblue", "6495ED" },
                { "steelblue", "4682B4" },
                { "royalblue", "4682B4" },
                { "blue", "0000FF" },
                { "mediumblue", "0000CD" },
                { "darkblue", "00008B" },
                { "navy", "000080" },
                { "midnightblue", "191970" },
                //Purple/Violet/Magenta colors
                { "lavender", "E6E6FA" },
                { "thistle", "D8BFD8" },
                { "plum", "DDA0DD" },
                { "violet", "EE82EE" },
                { "orchid", "DA70D6" },
                { "fuchsia", "FF00FF" },
                { "magenta", "FF00FF" },
                { "mediumorchid", "BA55D3" },
                { "mediumpurple", "9370DB" },
                { "blueviolet", "8A2BE2" },
                { "darkviolet", "9400D3" },
                { "darkorchid", "9932CC" },
                { "darkmagenta", "8B008B" },
                { "purple", "800080" },
                { "indigo", "4B0082" },
                { "darkslateblue", "483D8B" },
                { "rebeccapurple", "663399" },
                { "slateblue", "6A5ACD" },
                { "mediumslateblue", "7B68EE" },
                //White colors
                { "white", "FFFFFF" },
                { "snow", "FFFAFA" },
                { "honeydew", "F0FFF0" },
                { "mintcream", "F5FFFA" },
                { "azure", "F0FFFF" },
                { "aliceblue", "F0F8FF" },
                { "ghostwhite", "F8F8FF" },
                { "whitesmoke", "F5F5F5" },
                { "seashell", "FFF5EE" },
                { "beige", "F5F5DC" },
                { "oldlace", "FDF5E6" },
                { "floralwhite", "FFFAF0" },
                { "ivory", "FFFFF0" },
                { "antiqueWhite", "FAEBD7" },
                { "linen", "FAF0E6" },
                { "lavenderblush", "FFF0F5" },
                { "mistyrose", "FFE4E1" },
                //Gray/Black colors
                { "gainsboro", "DCDCDC" },
                { "lightgrey", "D3D3D3" },
                { "silver", "C0C0C0" },
                { "darkgray", "A9A9A9" },
                { "gray", "808080" },
                { "dimgray", "696969" },
                { "lightslategray", "778899" },
                { "slategray", "708090" },
                { "darkslategray", "2F4F4F" },
                { "black", "000000" }
            };
        }

        //Convert 3 char length hex to 6 char length hex
        private static string FillHex(string hex)
        {
            //Can apply this method. Thus return with original value
            if (hex.Length != 3)
            {
                return hex;
            }

            string _hex = string.Empty;

            for (int i = 0; 3 > i; i++)
            {
                _hex = string.Concat(_hex, hex[i], hex[i]);
            }

            return _hex;
        }

        internal static bool IsHex(string styleValue)
        {
            return styleValue.IndexOf("#") >= 0;
        }

        internal static string GetHexColor(string styleValue)
        {
            string hex = string.Empty;

            if (string.IsNullOrEmpty(styleValue))
            {
                return string.Empty;
            }

            if (IsRGB(styleValue))
            {
                hex = GetHex(styleValue);
            }
            else
                if (IsHex(styleValue))
            {
                hex = styleValue.Replace("#", string.Empty);

                if (!string.IsNullOrEmpty(hex) && hex.Length == 3)
                {
                    hex = FillHex(hex);
                }

                //Invalid hex value.
                if (hex.Length != 6)
                {
                    hex = string.Empty;
                }
            }
            else
            {
                if (namedColors.TryGetValue(styleValue.Trim().ToLower(), out string value))
                {
                    hex = value;
                }
            }

            return hex;
        }

        internal static void ApplyBackGroundColor(string style, OpenXmlElement styleElement)
        {
            string hex = GetHexColor(style);

            if (!string.IsNullOrEmpty(hex))
            {
                styleElement.Append(new Word.Shading()
                {
                    Fill = hex,
                    Color = "auto",
                    Val = Word.ShadingPatternValues.Clear
                });
            }
        }

        internal static void ApplyColor(string style, OpenXmlElement styleElement)
        {
            string hex = GetHexColor(style);

            if (!string.IsNullOrEmpty(hex))
            {
                styleElement.Append(new Word.Color() { Val = hex });
            }
        }

        internal static string ExtractBackGround(string style)
        {
            if (string.IsNullOrEmpty(style))
            {
                return string.Empty;
            }

            string[] splits = style.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            if (splits == null && splits.Length <= 0)
            {
                return string.Empty;
            }

            return splits[0];
        }
    }
}
