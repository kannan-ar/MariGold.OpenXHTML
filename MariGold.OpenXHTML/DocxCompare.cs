namespace MariGold.OpenXHTML
{
    using System;

    internal static class DocxCompare
    {
        internal static bool CompareStringOrdinalIgnoreCase(this string source, string value)
        {
            if(string.IsNullOrEmpty(source) || string.IsNullOrEmpty(value))
            {
                return false;
            }

            return string.Equals(source, value, StringComparison.OrdinalIgnoreCase);
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
