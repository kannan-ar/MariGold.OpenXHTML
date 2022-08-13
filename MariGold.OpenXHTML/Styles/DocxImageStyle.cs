namespace MariGold.OpenXHTML.Styles
{
    internal class DocxImageStyle
    {
        private readonly DocxNode node;

        internal const string widthName = "width";
        internal const string heightName = "height";

        internal DocxImageStyle(DocxNode node)
        {
            this.node = node;
        }

        internal void TryGetDimensions(out decimal? width, out decimal? height)
        {
            width = null;
            height = null;

            var widthStyle = node.ExtractStyleValue(widthName);
            var heightStyle = node.ExtractOwnStyleValue(heightName);

            if (DocxUnits.ConvertToPx(widthStyle, out decimal w))
            {
                width = w;
            }

            if (DocxUnits.ConvertToPx(heightStyle, out decimal h))
            {
                height = h;
            }
        }

        internal void ApplyInheritedStyles()
        {
            string widthStyleValue = node.ExtractOwnStyleValue(widthName);

            if (!string.IsNullOrEmpty(widthStyleValue)) return;

            string inheritedWidthValue = node.ExtractInheritedStyleValue(widthName);

            if (!string.IsNullOrEmpty(inheritedWidthValue))
            {
                node.SetExtentedStyle(widthName, inheritedWidthValue);
            }
        }

        internal decimal ScaleWithAspectRatio(decimal actualValue, decimal scaledValue, decimal toBeScaledValue)
        {
            return (actualValue / scaledValue) * toBeScaledValue;
        }
    }
}
