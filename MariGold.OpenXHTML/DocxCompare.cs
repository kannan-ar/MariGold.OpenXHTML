namespace MariGold.OpenXHTML
{
    using System;

    internal static class DocxCompare
    {
        internal static bool CompareStringInvariantCultureIgnoreCase(this string source, string value)
        {
            if(string.IsNullOrEmpty(source) || string.IsNullOrEmpty(value))
            {
                return false;
            }

            return string.Compare(source, value, StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal static bool HasStringContains(this string path, params string[] extensions)
        {
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }

            foreach(string ext in extensions)
            {
                if(path.Contains(ext))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
