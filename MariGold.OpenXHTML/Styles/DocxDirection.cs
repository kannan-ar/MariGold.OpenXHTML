using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace MariGold.OpenXHTML
{
    internal static class DocxDirection
    {
        internal const string direction = "direction";

        internal const string ltr = "ltr";
        internal const string rtl = "rtl";

        private static bool GetDirectionValue(string style, out DirectionValues direction)
        {
            direction = DirectionValues.Ltr;
            bool assigned = false;

            switch (style.ToLower())
            {
                case ltr:
                    assigned = true;
                    direction = DirectionValues.Ltr;
                    break;

                case rtl:
                    assigned = true;
                    direction = DirectionValues.Rtl;
                    break;
            }

            return assigned;
        }

        internal static void ApplyBidi(string style, OpenXmlElement styleElement)
        {
            if (GetDirectionValue(style, out DirectionValues direction) && direction == DirectionValues.Rtl)
            {
                styleElement.Append(new BiDi());
            }
        }

        internal static void ApplyDirection(string style, OpenXmlElement styleElement)
        {
            if (GetDirectionValue(style, out DirectionValues direction) && direction == DirectionValues.Rtl)
            {
                styleElement.Append(new RightToLeftText());
            }
        }
    }
}
