namespace MariGold.OpenXHTML.Styles
{
    internal static class DocxBlockStyle
    {
        private static string[] blockStyles =
        {
            "width"
        };

        internal static void ApplyBlockStyles(DocxNode parent, DocxNode child)
        {
            foreach(var blockStyle in blockStyles)
            {
                string value = child.ExtractStyleValue(blockStyle);

                if (!string.IsNullOrEmpty(value)) continue;

                value = parent.ExtractStyleValue(blockStyle);

                if (!string.IsNullOrEmpty(value))
                {
                    child.SetExtentedStyle(blockStyle, value);
                }
            }
        }
    }
}
