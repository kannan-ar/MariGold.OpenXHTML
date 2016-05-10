namespace MariGold.OpenXHTML
{
    using System;

    internal static class DocxCompare
    {
        public static bool CompareStringInvariantCultureIgnoreCase(this string source, string value)
        {
            if(string.IsNullOrEmpty(source) || string.IsNullOrEmpty(value))
            {
                return false;
            }

            return string.Compare(source, value, StringComparison.InvariantCultureIgnoreCase) == 0;
        }
    }
}
