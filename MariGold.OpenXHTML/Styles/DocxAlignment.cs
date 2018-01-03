namespace MariGold.OpenXHTML
{
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class DocxAlignment
	{
		internal const string textAlign = "text-align";
		internal const string verticalAlign = "vertical-align";

        internal const string center = "center";
        internal const string left = "left";
        internal const string right = "right";
	internal const string both = "justify";
        internal const string sub = "sub";
        internal const string super = "super";

        private static bool GetVerticalPositionValues(string style, out VerticalPositionValues alignment)
        {
            alignment = VerticalPositionValues.Baseline;
            bool assigned = false;

            switch (style.ToLower())
            {
                case sub:
                    assigned = true;
                    alignment = VerticalPositionValues.Subscript;
                    break;

                case super:
                    assigned = true;
                    alignment = VerticalPositionValues.Superscript;
                    break;
            }

            return assigned;
        }

        private static bool GetJustificationValue(string style, out JustificationValues alignment)
        {
            alignment = JustificationValues.Left;
            bool assigned = false;

            switch (style.ToLower())
            {
                case right:
                    assigned = true;
                    alignment = JustificationValues.Right;
                    break;

                case left:
                    assigned = true;
                    alignment = JustificationValues.Left;
                    break;

                case center:
                    assigned = true;
                    alignment = JustificationValues.Center;
                    break;
			    
	    	case both:
		    assigned = true;
		    alignment = JustificationValues.Both;
		    break;
            }

            return assigned;
        }
		
		internal static void ApplyTextAlign(string style, OpenXmlElement styleElement)
		{
			JustificationValues alignment;
					
			if (GetJustificationValue(style, out alignment))
			{
				styleElement.Append(new Justification() { Val = alignment });
			}
		}
		
		internal static bool GetCellVerticalAlignment(string style, out TableVerticalAlignmentValues alignment)
		{
			alignment = TableVerticalAlignmentValues.Top;
			bool assigned = false;

			switch (style.ToLower())
			{
				case "top":
					assigned = true;
					alignment = TableVerticalAlignmentValues.Top;
					break;

				case "middle":
					assigned = true;
					alignment = TableVerticalAlignmentValues.Center;
					break;

				case "bottom":
					assigned = true;
					alignment = TableVerticalAlignmentValues.Bottom;
					break;
			}

			return assigned;
		}
		
        internal static void ApplyVerticalTextAlign(string style, OpenXmlElement styleElement)
        {
            VerticalPositionValues alignment;

            if (GetVerticalPositionValues(style, out alignment))
            {
                styleElement.Append(new VerticalTextAlignment() { Val = alignment });
            }
        }
	}
}
