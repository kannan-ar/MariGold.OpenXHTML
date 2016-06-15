namespace MariGold.OpenXHTML
{
    using System;
    using System.Drawing;
    using DocumentFormat.OpenXml;
    using Word = DocumentFormat.OpenXml.Wordprocessing;
    using System.Collections.Generic;

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
                        int r, g, b = 0;

                        r = Convert.ToInt32(colors[0]);
                        g = Convert.ToInt32(colors[1]);
                        b = Convert.ToInt32(colors[2]);

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

            namedColors = new Dictionary<string, string>();

            namedColors.Add("pink", "FFC0CB");
            namedColors.Add("lightpink", "FFB6C1");
            namedColors.Add("hotpink", "FF69B4");
            namedColors.Add("deeppink", "FF1493");
            namedColors.Add("palevioletred", "DB7093");
            namedColors.Add("mediumvioletred", "C71585");
            namedColors.Add("lightsalmon", "FFA07A");
            namedColors.Add("salmon", "FA8072");
            namedColors.Add("darksalmon", "E9967A");
            namedColors.Add("lightcoral", "F08080");
            namedColors.Add("indianred", "CD5C5C");
            namedColors.Add("crimson", "DC143C");
            namedColors.Add("firebrick", "B22222");
            namedColors.Add("darkred", "8B0000");
            namedColors.Add("red", "FF0000");
            namedColors.Add("orangered", "FF4500");
            namedColors.Add("tomato", "FF6347");
            namedColors.Add("coral", "FF7F50");
            namedColors.Add("darkorange", "FF8C00");
            namedColors.Add("orange", "FFA500");
            namedColors.Add("yellow", "FFFF00");
            namedColors.Add("lightyellow", "FFFFE0");
            namedColors.Add("lemonchiffon", "FFFACD");
            namedColors.Add("lightgoldenrodyellow", "FAFAD2");
            namedColors.Add("papayawhip", "FFEFD5");
            namedColors.Add("moccasin", "FFE4B5");
            namedColors.Add("peachpuff", "FFDAB9");
            namedColors.Add("palegoldenrod", "EEE8AA");
            namedColors.Add("khaki", "F0E68C");
            namedColors.Add("darkkhaki", "BDB76B");
            namedColors.Add("gold", "FFD700");
            namedColors.Add("cornsilk", "FFF8DC");
            namedColors.Add("blanchedalmond", "FFEBCD");
            namedColors.Add("bisque", "FFE4C4");
            namedColors.Add("navajowhite", "FFDEAD");
            namedColors.Add("wheat", "F5DEB3");
            namedColors.Add("burlywood", "DEB887");
            namedColors.Add("tan", "D2B48C");
            namedColors.Add("rosybrown", "BC8F8F");
            namedColors.Add("sandybrown", "F4A460");
            namedColors.Add("goldenrod", "DAA520");
            namedColors.Add("darkgoldenrod", "B8860B");
            namedColors.Add("peru", "CD853F");
            namedColors.Add("chocolate", "D2691E");
            namedColors.Add("saddlebrown", "8B4513");
            namedColors.Add("sienna", "A0522D");
            namedColors.Add("brown", "A52A2A");
            namedColors.Add("maroon", "800000");
            namedColors.Add("darkolivegreen", "556B2F");
            namedColors.Add("olive", "808000");
            namedColors.Add("olivedrab", "6B8E23");
            namedColors.Add("yellowgreen", "9ACD32");
            namedColors.Add("limegreen", "32CD32");
            namedColors.Add("lime", "00FF00");
            namedColors.Add("lawngreen", "7CFC00");
            namedColors.Add("chartreuse", "7FFF00");
            namedColors.Add("greenyellow", "ADFF2F");
            namedColors.Add("springgreen", "00FF7F");
            namedColors.Add("mediumspringgreen", "00FA9A");
            namedColors.Add("lightgreen", "90EE90");
            namedColors.Add("palegreen", "98FB98");
            namedColors.Add("darkseagreen", "8FBC8F");
            namedColors.Add("mediumaquamarine", "66CDAA");
            namedColors.Add("mediumseagreen", "3CB371");
            namedColors.Add("seagreen", "2E8B57");
            namedColors.Add("forestgreen", "228B22");
            namedColors.Add("green", "008000");
            namedColors.Add("darkgreen", "006400");
            //Cyan colors
            namedColors.Add("aqua", "00FFFF");
            namedColors.Add("cyan", "00FFFF");
            namedColors.Add("lightcyan", "E0FFFF");
            namedColors.Add("paleturquoise", "AFEEEE");
            namedColors.Add("aquamarine", "7FFFD4");
            namedColors.Add("turquoise", "40E0D0");
            namedColors.Add("mediumturquoise", "48D1CC");
            namedColors.Add("darkturquoise", "00CED1");
            namedColors.Add("lightseagreen", "20B2AA");
            namedColors.Add("cadetblue", "5F9EA0");
            namedColors.Add("darkcyan", "008B8B");
            namedColors.Add("teal", "008080");
            //Blue colors
            namedColors.Add("lightsteelblue", "B0C4DE");
            namedColors.Add("powderblue", "B0E0E6");
            namedColors.Add("lightblue", "ADD8E6");
            namedColors.Add("skyblue", "87CEEB");
            namedColors.Add("lightskyblue", "87CEFA");
            namedColors.Add("deepskyblue", "00BFFF");
            namedColors.Add("dodgerblue", "1E90FF");
            namedColors.Add("cornflowerblue", "6495ED");
            namedColors.Add("steelblue", "4682B4");
            namedColors.Add("royalblue", "4682B4");
            namedColors.Add("blue", "0000FF");
            namedColors.Add("mediumblue", "0000CD");
            namedColors.Add("darkblue", "00008B");
            namedColors.Add("navy", "000080");
            namedColors.Add("midnightblue", "191970");
            //Purple/Violet/Magenta colors
            namedColors.Add("lavender", "E6E6FA");
            namedColors.Add("thistle", "D8BFD8");
            namedColors.Add("plum", "DDA0DD");
            namedColors.Add("violet", "EE82EE");
            namedColors.Add("orchid", "DA70D6");
            namedColors.Add("fuchsia", "FF00FF");
            namedColors.Add("magenta", "FF00FF");
            namedColors.Add("mediumorchid", "BA55D3");
            namedColors.Add("mediumpurple", "9370DB");
            namedColors.Add("blueviolet", "8A2BE2");
            namedColors.Add("darkviolet", "9400D3");
            namedColors.Add("darkorchid", "9932CC");
            namedColors.Add("darkmagenta", "8B008B");
            namedColors.Add("purple", "800080");
            namedColors.Add("indigo", "4B0082");
            namedColors.Add("darkslateblue", "483D8B");
            namedColors.Add("rebeccapurple", "663399");
            namedColors.Add("slateblue", "6A5ACD");
            namedColors.Add("mediumslateblue", "7B68EE");
            //White colors
            namedColors.Add("white", "FFFFFF");
            namedColors.Add("snow", "FFFAFA");
            namedColors.Add("honeydew", "F0FFF0");
            namedColors.Add("mintcream", "F5FFFA");
            namedColors.Add("azure", "F0FFFF");
            namedColors.Add("aliceblue", "F0F8FF");
            namedColors.Add("ghostwhite", "F8F8FF");
            namedColors.Add("whitesmoke", "F5F5F5");
            namedColors.Add("seashell", "FFF5EE");
            namedColors.Add("beige", "F5F5DC");
            namedColors.Add("oldlace", "FDF5E6");
            namedColors.Add("floralwhite", "FFFAF0");
            namedColors.Add("ivory", "FFFFF0");
            namedColors.Add("antiqueWhite", "FAEBD7");
            namedColors.Add("linen", "FAF0E6");
            namedColors.Add("lavenderblush", "FFF0F5");
            namedColors.Add("mistyrose", "FFE4E1");
            //Gray/Black colors
            namedColors.Add("gainsboro", "DCDCDC");
            namedColors.Add("lightgrey", "D3D3D3");
            namedColors.Add("silver", "C0C0C0");
            namedColors.Add("darkgray", "A9A9A9");
            namedColors.Add("gray", "808080");
            namedColors.Add("dimgray", "696969");
            namedColors.Add("lightslategray", "778899");
            namedColors.Add("slategray", "708090");
            namedColors.Add("darkslategray", "2F4F4F");
            namedColors.Add("black", "000000");
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
                    string value;

                    if (namedColors.TryGetValue(styleValue.Trim().ToLower(), out value))
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
                    Color = hex,
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
